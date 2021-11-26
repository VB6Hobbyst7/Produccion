VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogReqPresupu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Cuenta Contable del Presupuesto"
   ClientHeight    =   6165
   ClientLeft      =   1230
   ClientTop       =   1950
   ClientWidth     =   8610
   Icon            =   "frmLogReqPresupu.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboPreRub 
      Height          =   315
      Left            =   1365
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   795
      Width           =   4050
   End
   Begin VB.TextBox txtTipCam 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   6870
      TabIndex        =   13
      Top             =   510
      Width           =   915
   End
   Begin VB.ComboBox cboPrePri 
      Height          =   315
      Left            =   1365
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   435
      Width           =   4050
   End
   Begin VB.TextBox txtCtaCnt 
      Height          =   315
      Left            =   6885
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.TextBox txtNomPro 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1365
      TabIndex        =   5
      Top             =   60
      Width           =   6990
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4005
      TabIndex        =   4
      Top             =   5580
      Width           =   1140
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Deshacer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5490
      TabIndex        =   3
      Top             =   5580
      Width           =   1140
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6945
      TabIndex        =   2
      Top             =   5580
      Width           =   1140
   End
   Begin VB.ComboBox cboPreRubCta 
      Height          =   315
      Left            =   1365
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1155
      Width           =   2490
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgPresu 
      Height          =   3885
      Left            =   225
      TabIndex        =   0
      Top             =   1545
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   6853
      _Version        =   393216
      Cols            =   4
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483638
      FocusRect       =   0
      HighLight       =   2
      FillStyle       =   1
      RowSizingMode   =   1
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Tipo Cambio :"
      Height          =   255
      Index           =   4
      Left            =   5745
      TabIndex        =   12
      Top             =   570
      Width           =   1020
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Rubro :"
      Height          =   225
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   870
      Width           =   930
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Presupuesto :"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   495
      Width           =   1080
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Cta.Contable :"
      Height          =   210
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   1215
      Width           =   1065
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Item :"
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   105
      Width           =   750
   End
End
Attribute VB_Name = "frmLogReqPresupu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pRpta As String
Dim pFrmTpo As String, pPeriodo As String, pReqNro As String, pReqTraNro As String, _
    pBSCod As String, pCtaCont As String
Dim pPresu As String, pAno As String, pMone As String
Dim pPreUni As Currency
Dim bIniCta As Boolean, bIniPre As Boolean
Dim lsCtaContCod As String
Dim lbOk As Boolean

Public Function Inicio(ByVal psFrmTpo As String, ByVal psPeriodo As String, ByVal psReqNro As String, _
ByVal psReqTraNro As String, ByVal psBsCod As String, ByVal psCtaContCod As String, _
Optional bLectura As Boolean = False) As String
On Error GoTo ErrorInicia
bIniCta = False
bIniPre = False

pFrmTpo = psFrmTpo
pPeriodo = psPeriodo
pReqNro = psReqNro
pReqTraNro = psReqTraNro
pBSCod = psBsCod
pCtaCont = psCtaContCod

If bLectura Then
    cmdGrabar.Visible = False
    cmdModificar.Visible = False
    cboPrePri.Enabled = False
    cboPreRubCta.Enabled = False
End If
Me.Show 1
Inicio = pRpta
Exit Function

ErrorInicia:
    MsgBox "Error Nº [" & Str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
    Screen.MousePointer = 0
End Function


Private Sub cboPreRub_Click()
    Dim clsDPre As DPresupu
    Dim rs As ADODB.Recordset
    
    Set clsDPre = New DPresupu
    Set rs = New ADODB.Recordset
    
    'Carga Ctas
    Set rs = clsDPre.CargaPreRubCta(pPeriodo, Right(Trim(cboPrePri.Text), 1), Trim(Right(Trim(cboPreRub.Text), 30)))
    CargaCombo cboPreRubCta, rs
    
    Set rs = Nothing
    Set clsDPre = Nothing
End Sub

Private Sub Form_Load()
Dim clsDReq As DLogRequeri
Dim clsDPre As DPresupu
Dim rs As ADODB.Recordset
Dim tmpReg As New ADODB.Recordset
Dim tmpSql As String
Dim M As Integer

Set clsDReq = New DLogRequeri
Set clsDPre = New DPresupu
Set rs = New ADODB.Recordset

Call CentraForm(Me)

Set rs = clsDReq.CargaReqDetalle(ReqDetUnRegistroPresupu, pReqNro, pReqTraNro, pBSCod)
With rs
    If .RecordCount = 1 Then
        'pPresu = !cPresu
        'pAno = !cAno
        'pMone = !cLogReqMoneda
        'pPreUni = !nLogReqDetPrecio
        'txtCtaCnt = !cCtaContCod
        txtTipCam = !nLogReqTipCambio
        txtNomPro = !cBSDescripcion
        pRpta = txtCtaCnt
        'If tmpReg!cEstFoc = "0" Then
        '    cmdModificar.Enabled = False
        'End If
    End If
End With
Set clsDReq = Nothing

'Carga Presupuestos
Set rs = clsDPre.CargaPrePrimario(PrePriCbo)
CargaCombo cboPrePri, rs

Set rs = Nothing
Set clsDPre = Nothing

'Carga las ctas.cnts.
''tmpSql = " Select Distinct substring(cCtaCnt,1,2) + '1' + substring(cCtaCnt,4,len(cCtaCnt)-3)  Campo1, substring(cCtaCnt,1,2) + '1' + substring(cCtaCnt,4,len(cCtaCnt)-3) Campo2 From PRubCta Where cPresu = '" & Right(Trim(cboPrePri), 4) & "'"
''Call CboBox(cboCta, tmpSql, " ", txtCtaCnt)

''If pMone = "2" Then
''    lblEtiqueta(4).Visible = True
''    txtTipCam.Visible = True
''    If Len(Trim(txtCtaCnt)) = 0 Then
''        txtTipCam = Format(Val(FuncGnral("Select nValFijoDia Campo From tipcambio Where datediff(dd, dfeccamb, '" & Format(gdFecSis, "mm/dd/yyyy") & "') = 0 ")), "#0.000")
''    End If
''End If
''If Len(pPresu) > 0 Then
''    Call CargaPreCta(txtCtaCnt)
''    cboPrePri.Enabled = False
''    cboCta.Enabled = False
''    cmdGrabar.Enabled = False
''    fgPresu.Col = 9
''    For M = 1 To fgPresu.Rows - 1
''        fgPresu.Row = M
''        fgPresu.CellBackColor = &H80000018     ' color grabados
''    Next
''Else
''    cmdModificar.Enabled = False
''End If
bIniCta = True
bIniPre = True
End Sub

Private Sub cboCta_Click()
If bIniCta = True Then
    txtCtaCnt = Trim(Left(cboCta.Text, 50))
    If VeriCta(txtCtaCnt) = True Then
        Call CargaPreCta(txtCtaCnt)
    End If
End If
End Sub

Private Sub cboPrePri_Click()
    Dim clsDPre As DPresupu
    Dim rs As ADODB.Recordset
    
    Set clsDPre = New DPresupu
    Set rs = New ADODB.Recordset
    
    'Carga Rubros
    Set rs = clsDPre.CargaPreRubro(PreRubCbo, pPeriodo, Right(Trim(cboPrePri.Text), 1))
    CargaCombo cboPreRub, rs
    
    Set rs = Nothing
    Set clsDPre = Nothing
'    Call CboBox(cboCta, "Select Distinct substring(cCtaCnt,1,2) + '1' + substring(cCtaCnt,4,len(cCtaCnt)-3)  Campo1, cCtaCnt Campo2 From PRubCta Where cPresu = '" & Right(Trim(cboPrePri), 4) & "'")
'    txtCtaCnt.Text = ""
'    txtNomRub.Text = ""
End Sub

Private Sub cmdGrabar_Click()
Dim tmpSql As String
Dim M As Integer, N As Integer
Dim vAcumu As Currency
pPresu = Right(Trim(cboPrePri), 4)
If Len(pPresu) = 0 Then
    MsgBox "Presupuesto no determinado", vbInformation, " Aviso "
    Exit Sub
End If
If Len(Trim(txtCtaCnt)) <= 0 Then
    MsgBox "Cuenta contable no ingresada", vbInformation, " Aviso "
    Exit Sub
End If
If VeriCta(txtCtaCnt) = False Then
    Exit Sub
End If
vAcumu = 0
For N = 1 To fgPresu.Rows - 1
    vAcumu = vAcumu + (CCur(fgPresu.TextMatrix(N, 6)) - (CCur(fgPresu.TextMatrix(N, 8)) + CCur(fgPresu.TextMatrix(N, 9))))
    'If CCur(fgPresu.TextMatrix(N, 6)) < CCur(fgPresu.TextMatrix(N, 9)) Then
    If vAcumu < CCur(fgPresu.TextMatrix(N, 9)) Then
        If vAcumu < CCur(fgPresu.TextMatrix(N, 9)) Then
            MsgBox "Monto a reservar es mayor al presupuesto" & vbCr & "en el periodo de " & fgPresu.TextMatrix(N, 2), vbInformation, "Aviso"
            Exit Sub
        Else
            If MsgBox("Monto a reservar es mayor al presupuesto" & vbCr & "en el periodo de " & fgPresu.TextMatrix(N, 2) & vbCr & "Desea contiuar ya que en el acumulado todavia es inferior ?", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                vAcumu = vAcumu - CCur(fgPresu.TextMatrix(N, 9))
            Else
                Exit Sub
            End If
        End If
    End If
Next
If MsgBox("Esta seguro que desea reservar en esta cuenta ?", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
    dbCmact.BeginTrans
    'Actualiza la cta cnt, presu, tipcambio en LObtenDet
    tmpSql = "Update LObtenDet Set cCtaCnt = '" & Trim(txtCtaCnt) & "', cPresu = '" & pPresu & "', nTipCam = " & Val(txtTipCam) & "  Where cObjetoCod = '" & pBSCod & "'And cNroObt = '" & pReqNro & "'"
    dbCmact.Execute tmpSql
    
    'Reserva monto en PPresupu
    For M = 1 To fgPresu.Rows - 1
        If CCur(fgPresu.TextMatrix(M, 9)) > 0 Then
            tmpSql = "Update PPresupu Set nMonRes =  nMonRes +  " & CCur(fgPresu.TextMatrix(M, 9)) & " " & _
                " Where cAno = '" & pAno & "' And cPresu = '" & pPresu & "' " & _
                " And cCodRub = '" & fgPresu.TextMatrix(M, 3) & "' And cPeriodo = '" & fgPresu.TextMatrix(M, 1) & "'" & _
                " "
            dbCmact.Execute tmpSql
        End If
    Next
    dbCmact.CommitTrans

    cboPrePri.Enabled = False
    cboCta.Enabled = False
    cmdGrabar.Enabled = False
    cmdModificar.Enabled = True
    pRpta = Trim(txtCtaCnt)
    
    Call CargaPreCta(txtCtaCnt)
    
    fgPresu.Col = 9
    For M = 1 To fgPresu.Rows - 1
        fgPresu.Row = M
        fgPresu.CellBackColor = &H80000018     ' color de los grabados
    Next
End If
End Sub

Private Sub CmdModificar_Click()
Dim tmpSql As String
Dim M As Integer
If MsgBox("Esta seguro que desea eliminar la reserva de esta Cuenta ?", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
    
    dbCmact.BeginTrans
    'Actualiza la cta cnt, presu, tipcambio en LObtenDet
    tmpSql = "Update LObtenDet Set cCtaCnt = Null , cPresu = Null , nTipCam = Null  Where cObjetoCod = '" & pBSCod & "'And cNroObt = '" & pReqNro & "'"
    dbCmact.Execute tmpSql
    
    'Resta monto en Reserva de PPresupu
    For M = 1 To fgPresu.Rows - 1
        If CCur(fgPresu.TextMatrix(M, 9)) > 0 Then
            tmpSql = "Update PPresupu Set nMonRes =  nMonRes -  " & CCur(fgPresu.TextMatrix(M, 9)) & " " & _
                " Where cAno = '" & pAno & "' And cPresu = '" & pPresu & "' " & _
                " And cCodRub = '" & fgPresu.TextMatrix(M, 3) & "' And cPeriodo = '" & fgPresu.TextMatrix(M, 1) & "'" & _
                " "
            dbCmact.Execute tmpSql
        End If
    Next
    dbCmact.CommitTrans
    
    pRpta = ""
    cboPrePri.Enabled = True
    cboCta.Enabled = True
    cmdGrabar.Enabled = True
    txtTipCam = Format(Val(FuncGnral("Select nValFijoDia Campo From tipcambio Where datediff(dd, dfeccamb, '" & Format(gdFecSis, "mm/dd/yyyy") & "') = 0 ")), "#0.000")
    cmdModificar.Enabled = False
    
    Call CargaPreCta(txtCtaCnt)
    Ok = True
    lsCtaContCod = ""
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub txtCtaCnt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If VeriCta(txtCtaCnt) = True Then
        Call CargaPreCta(txtCtaCnt)
    End If
End If
End Sub

Private Function VeriCta(ByVal cCta As String) As Boolean
Dim tmpSql As String
Dim vCant As Integer
VeriCta = False
'Se verifica si la cta. cnt. esta la Tabla CtaCont
vCant = Val(FuncGnral("Select  Count(*) Campo From CtaCont Where cCtaContCod like '" & cCta & "%'"))
If vCant = 1 Then
    VeriCta = True
ElseIf vCant = 0 Then
    MsgBox "Cuenta no existe", vbInformation, "Aviso"
ElseIf vCant > 1 Then
    MsgBox "Cuenta no esta en el Ultimo Nivel", vbInformation, "Aviso"
End If
End Function

Private Sub CargaPreCta(ByVal sCtaCnt As String)
Dim tmpReg As New ADODB.Recordset
Dim tmpSql As String
Dim x As Integer
Dim vAcumu As Currency
Call MSHFlex(fgPresu, 10, "Item-cCodPer-Periodo-Rubro Pres.-Inicial-Ejecutado-Actual-Act.Acumul.-Reservado-Por Reserv.", "450-0-900-0-1100-1100-1100-1100-1100-1100", "C-L-L-L-R-R-R-R-R-R")
tmpSql = " SELECT t.cCodTab, t.cNomTab, b.cCodRub, b.cDesRub, p.nMonIni, p.nMonAct, p.nMonEje, p.nMonRes " & _
    " FROM TablaCod T Left Join PPresupu P ON t.ccodtab = p.cPeriodo  " & _
    " Inner Join PRubCta R On p.cAno = r.cAno and p.cPresu = r.cPresu and p.cCodRub = r.cCodRub " & _
    " AND r.cAno = '" & pAno & "' and  p.cmoneda = '1' inner join PRubro B On r.ccodrub=b.ccodrub" & _
    " Where t.cCodTab like 'P3__' AND t.cCodTab <> 'P300' and p.cPresu = '" & Right(Trim(cboPrePri), 4) & "' " & _
    " and substring(cctacnt,1,2) = '" & Left(sCtaCnt, 2) & "' and substring(cctacnt,4,len(cctacnt)-3) = '" & Mid(sCtaCnt, 4) & "' Order By b.cdesrub, t.cCodTab "
tmpReg.Open tmpSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
If (tmpReg.BOF Or tmpReg.EOF) Then
Else
    With tmpReg
        txtNomRub = !cDesRub
        Do While Not tmpReg.EOF
            x = x + 1
            vAcumu = vAcumu + Val(!nMonAct) - Val(!nMonRes)
            AdicionaRow fgPresu, x
            fgPresu.Row = fgPresu.Rows - 1
            fgPresu.TextMatrix(x, 0) = x
            fgPresu.TextMatrix(x, 1) = !cCodTab
            fgPresu.TextMatrix(x, 2) = Trim(!cNomTab)
            fgPresu.TextMatrix(x, 3) = Trim(!cCodRub)
            fgPresu.TextMatrix(x, 4) = Format(!nMonIni, "#,##0.00")
            fgPresu.TextMatrix(x, 5) = Format(!nMonEje, "#,##0.00")
            fgPresu.TextMatrix(x, 6) = Format(!nMonAct, "#,##0.00")
            fgPresu.TextMatrix(x, 7) = Format(vAcumu, "#,##0.00")
            fgPresu.TextMatrix(x, 8) = Format(!nMonRes, "#,##0.00")
            If pMone = "1" Then
                fgPresu.TextMatrix(x, 9) = Format(FuncGnral("Select (p.nCantApro * d.nMonPre) Campo From LObtenDet D Inner Join LObtenPer P On d.cNroObt = p.cnroobt and d.cobjetocod = p.cobjetocod Where p.cNroObt = '" & pReqNro & "' and p.cObjetocod = '" & pBSCod & "' and p.cPeriodo = '" & !cCodTab & "'"), "#,##0.00")
            Else
                fgPresu.TextMatrix(x, 9) = Format(FuncGnral("Select (p.nCantApro * d.nMonPre * " & Val(txtTipCam) & ") campo From LObtenDet D Inner Join LObtenPer P On d.cNroObt = p.cnroobt and d.cobjetocod = p.cobjetocod Where p.cNroObt = '" & pReqNro & "' and p.cObjetocod = '" & pBSCod & "' and p.cPeriodo = '" & !cCodTab & "'"), "#,##0.00")
            End If
            'If CCur(fgPresu.TextMatrix(X, 9)) > 0 Then
            '    fgPresu.TextMatrix(X, 8) = Format(CCur(fgPresu.TextMatrix(X, 8)) - CCur(fgPresu.TextMatrix(X, 9)), "#,##0.00")
            'End If
            .MoveNext
        Loop
    End With
End If
tmpReg.Close
Set tmpReg = Nothing
End Sub
Public Property Get Ok() As Boolean
Ok = lbOk
End Property
Public Property Let Ok(ByVal vNewValue As Boolean)
lbOk = vNewValue
End Property
Private Property Get CtaContCod() As String
CtaContCod = lsCtaContCod
End Property
Private Property Let CtaContCod(ByVal vNewValue As String)
lsCtaContCod = vNewValue
End Property
