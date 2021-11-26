VERSION 5.00
Begin VB.Form frmLogReqPresupu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Cuenta Contable del Presupuesto"
   ClientHeight    =   6060
   ClientLeft      =   795
   ClientTop       =   1920
   ClientWidth     =   7575
   Icon            =   "frmLogReqPresupu.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Desasignar"
      Enabled         =   0   'False
      Height          =   390
      Left            =   3480
      TabIndex        =   14
      Top             =   1575
      Width           =   1305
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Asignar"
      Enabled         =   0   'False
      Height          =   390
      Left            =   1485
      TabIndex        =   13
      Top             =   1575
      Width           =   1305
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   5475
      TabIndex        =   1
      Top             =   1575
      Width           =   1305
   End
   Begin VB.ComboBox cboPreRub 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1230
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   765
      Width           =   4455
   End
   Begin VB.ComboBox cboPrePri 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1230
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   405
      Width           =   4455
   End
   Begin VB.TextBox txtNomPro 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1230
      TabIndex        =   4
      Top             =   60
      Width           =   6090
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
      Left            =   90
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
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
      Left            =   90
      TabIndex        =   2
      Top             =   1635
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ComboBox cboPreRubCta 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1230
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1125
      Width           =   3090
   End
   Begin Sicmact.FlexEdit fgeMes 
      Height          =   3870
      Left            =   135
      TabIndex        =   15
      Top             =   2055
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   6826
      Cols0           =   7
      HighLight       =   1
      AllowUserResizing=   1
      EncabezadosNombres=   "Item-Código-Mes-Inicial-Ejecutado-Resuelto-Por Asignar"
      EncabezadosAnchos=   "450-0-1100-1200-1200-1200-1200"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "R-L-L-R-R-R-R"
      FormatosEdit    =   "0-0-0-2-2-2-2"
      CantEntero      =   6
      AvanceCeldas    =   1
      TextArray0      =   "Item"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   2
      lbFormatoCol    =   -1  'True
      lbPuntero       =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      Appearance      =   0
      ColWidth0       =   450
      RowHeight0      =   285
   End
   Begin VB.Label lblPeriodo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   6585
      TabIndex        =   12
      Top             =   435
      Width           =   735
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Periodo :"
      Height          =   255
      Index           =   5
      Left            =   5910
      TabIndex        =   11
      Top             =   450
      Width           =   690
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Rubro :"
      Height          =   225
      Index           =   3
      Left            =   165
      TabIndex        =   9
      Top             =   870
      Width           =   930
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Presupuesto :"
      Height          =   255
      Index           =   2
      Left            =   165
      TabIndex        =   7
      Top             =   495
      Width           =   1080
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Cta.Contable :"
      Height          =   210
      Index           =   1
      Left            =   165
      TabIndex        =   6
      Top             =   1215
      Width           =   1065
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Item :"
      Height          =   240
      Index           =   0
      Left            =   165
      TabIndex        =   5
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
    pBSCod As String
Dim pCtaContCod As String, pCtaContDat As String
Dim pPresu As String, pAno As String, pMone As String
Dim bIniCboRub As Boolean
'Dim pRS As ADODB.Recordset
Dim pTipCambio As Currency
Dim pPreUni As Currency

Public Function Inicio(ByVal psFrmTpo As String, ByVal psPeriodo As String, _
ByVal psReqNro As String, ByVal psReqTraNro As String, ByVal psBSCod As String, _
ByVal psCtaCont As String, ByVal pnTipCambio As Currency, Optional bLectura As Boolean = False) As String
On Error GoTo ErrorInicia
bIniCboRub = False

pFrmTpo = psFrmTpo
pPeriodo = psPeriodo
pReqNro = psReqNro
pReqTraNro = psReqTraNro
pBSCod = psBSCod
pTipCambio = pnTipCambio

pRpta = psCtaCont
If Trim(pRpta) <> "" Then
    pCtaContCod = Trim(Left(psCtaCont, 40))
    pCtaContDat = Trim(Mid(psCtaCont, 40))
Else
    pCtaContCod = ""
    pCtaContDat = ""
End If
'Si ya hay información


'Solo para mostrar los datos
If bLectura Then
    cmdGrabar.Visible = False
    cmdModificar.Visible = False
    CmdAceptar.Enabled = False
    cmdCancelar.Enabled = False
    cboPrePri.Enabled = False
    cboPreRub.Enabled = False
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
    Dim nPos As Integer, nCont As Integer
    If bIniCboRub Then
        Set clsDPre = New DPresupu
        Set rs = New ADODB.Recordset
        
        'Carga Distribución en Meses
        nPos = InStr(1, cboPreRub.Text, " ", vbTextCompare)
        Set rs = clsDPre.CargaPlaRubMes(PlaRubMesPreReq, Right(Trim(cboPrePri.Text), 6), Trim(Left(cboPreRub.Text, nPos + 1)), _
            lblPeriodo.Caption, pReqNro, pReqTraNro, pBSCod, pTipCambio)
        If rs.RecordCount > 0 Then
            Set fgeMes.Recordset = rs
            fgeMes.Col = 6
            For nCont = 1 To fgeMes.Rows - 1
                fgeMes.Row = nCont
                fgeMes.CellBackColor = &H80000018
            Next
        Else
            fgeMes.Clear
            fgeMes.FormaCabecera
            fgeMes.Rows = 2
        End If
        
        'Carga Ctas
        Set rs = clsDPre.CargaPlaRubCta(Right(Trim(cboPrePri.Text), 6), Trim(Left(cboPreRub.Text, nPos + 1)), Me.lblPeriodo.Caption)
        CargaCombo rs, cboPreRubCta
        
        Set rs = Nothing
        Set clsDPre = Nothing
        CmdAceptar.Enabled = False
    End If
End Sub

Private Sub cboPreRubCta_Click()
    If Trim(cboPreRubCta.Text) <> "" Then
        CmdAceptar.Enabled = True
    End If
End Sub

Private Sub CmdAceptar_Click()
Dim nPos As Integer
If Trim(cboPreRubCta.Text) = "" Then
    MsgBox "No se ha determinado la Cuenta a asignar", vbInformation, " Aviso "
    Exit Sub
End If
If fgeMes.Rows = 2 Then
    MsgBox "No existe partida en el Presupuesto", vbInformation, " Aviso "
    Exit Sub
End If


nPos = InStr(1, cboPreRub.Text, " ", vbTextCompare)
pRpta = Trim(cboPreRubCta.Text) & Space(40) & Right(Trim(cboPrePri.Text), 6) & "-" & Trim(Left(cboPreRub.Text, nPos + 1))
Unload Me
End Sub

Private Sub CmdCancelar_Click()
pRpta = "   "
Unload Me
End Sub

Private Sub Form_Load()
Dim clsDReq As DLogRequeri
Dim clsDPre As DPresupu
Dim rs As ADODB.Recordset
Dim tmpReg As New ADODB.Recordset
Dim tmpSql As String
Dim m As Integer
Dim nCont As Integer, nPos As Integer

Set clsDReq = New DLogRequeri
Set clsDPre = New DPresupu
Set rs = New ADODB.Recordset

Call CentraForm(Me)

Set rs = clsDReq.CargaReqDetalle(ReqDetUnRegistroPresupu, pReqNro, pReqTraNro, pBSCod)
With rs
    If .RecordCount = 1 Then
        txtNomPro = !cBSDescripcion
        pPreUni = !nLogReqDetPrecio
        
        'Set pRS = New ADODB.Recordset
        'Set pRS = clsDReq.CargaReqDetMes(ReqDetMesUnRegistro, pReqNro, pReqTraNro, pBSCod)
    End If
End With
Set clsDReq = Nothing

lblPeriodo.Caption = pPeriodo
CambiaTamañoCombo cboPrePri, 400
CambiaTamañoCombo cboPreRub, 400

'Carga Presupuestos
Set rs = clsDPre.CargaPlaPresupu(PlaPriCbo, pPeriodo)
CargaCombo rs, cboPrePri

If Trim(pCtaContDat) <> "" Then
    For nCont = 0 To cboPrePri.ListCount - 1
        cboPrePri.ListIndex = nCont
        If Right(Trim(cboPrePri.Text), 6) = Trim(Left(pCtaContDat, 6)) Then Exit For
    Next
    
    'Carga Rubros
    Set rs = clsDPre.CargaPlaRubro(PlaRubCbo, Right(Trim(cboPrePri.Text), 6), Me.lblPeriodo.Caption)
    CargaCombo rs, cboPreRub
    
    For nCont = 0 To cboPreRub.ListCount - 1
        cboPreRub.ListIndex = nCont
        nPos = InStr(1, cboPreRub.Text, " ", vbTextCompare)
        If Trim(Left(cboPreRub.Text, nPos + 1)) = Trim(Mid(pCtaContDat, 8)) Then Exit For
    Next
    
    'Carga Ctas
    nPos = InStr(1, cboPreRub.Text, " ", vbTextCompare)
    Set rs = clsDPre.CargaPlaRubCta(Right(Trim(cboPrePri.Text), 6), Trim(Left(cboPreRub.Text, nPos + 1)), Me.lblPeriodo.Caption)
    CargaCombo rs, cboPreRubCta
    
    UbicaCombo cboPreRubCta, pCtaContCod, False, Len(pCtaContCod)
    
    'Carga Distribución en Meses
    Set rs = clsDPre.CargaPlaRubMes(PlaRubMesPreReq, Right(Trim(cboPrePri.Text), 6), Trim(Left(cboPreRub.Text, nPos + 1)), _
        pReqNro, pReqTraNro, pBSCod, pTipCambio)
    If rs.RecordCount > 0 Then
        Set fgeMes.Recordset = rs
        
        fgeMes.Col = 6
        For nCont = 1 To fgeMes.Rows - 1
            fgeMes.Row = nCont
            fgeMes.CellBackColor = &H80000018
        Next
    Else
        fgeMes.Clear
        fgeMes.FormaCabecera
        fgeMes.Rows = 2
    End If
    
    CmdAceptar.Enabled = True
    cmdCancelar.Enabled = True
End If
Set rs = Nothing
Set clsDPre = Nothing

bIniCboRub = True
End Sub

'''Private Sub cboCta_Click()
'''If bIniCta = True Then
'''    txtCtaCnt = Trim(Left(cboCta.Text, 50))
'''    If VeriCta(txtCtaCnt) = True Then
'''        Call CargaPreCta(txtCtaCnt)
'''    End If
'''End If
'''End Sub

Private Sub cboPrePri_Click()
    Dim clsDPre As DPresupu
    Dim rs As ADODB.Recordset
    
    If bIniCboRub Then
        Set clsDPre = New DPresupu
        Set rs = New ADODB.Recordset
        
        'Carga Rubros
        Set rs = clsDPre.CargaPlaRubro(PlaRubCbo, Right(Trim(cboPrePri.Text), 6), lblPeriodo.Caption)
        CargaCombo rs, cboPreRub
        
        Set rs = Nothing
        Set clsDPre = Nothing
        cboPreRubCta.Clear
        
        CmdAceptar.Enabled = False
    End If
'    Call CboBox(cboCta, "Select Distinct substring(cCtaCnt,1,2) + '1' + substring(cCtaCnt,4,len(cCtaCnt)-3)  Campo1, cCtaCnt Campo2 From PRubCta Where cPresu = '" & Right(Trim(cboPrePri), 4) & "'")
'    txtCtaCnt.Text = ""
'    txtNomRub.Text = ""
End Sub

'''Private Sub cmdGrabar_Click()
'''Dim tmpSql As String
'''Dim m As Integer, n As Integer
'''Dim vAcumu As Currency
'''pPresu = Right(Trim(cboPrePri), 4)
'''If Len(pPresu) = 0 Then
'''    MsgBox "Presupuesto no determinado", vbInformation, " Aviso "
'''    Exit Sub
'''End If
'''If Len(Trim(cboCtaCnt)) <= 0 Then
'''    MsgBox "Cuenta contable no ingresada", vbInformation, " Aviso "
'''    Exit Sub
'''End If
'''If VeriCta(txtCtaCnt) = False Then
'''    Exit Sub
'''End If
'''vAcumu = 0
'''For n = 1 To fgPresu.Rows - 1
'''    vAcumu = vAcumu + (CCur(fgPresu.TextMatrix(n, 6)) - (CCur(fgPresu.TextMatrix(n, 8)) + CCur(fgPresu.TextMatrix(n, 9))))
'''    'If CCur(fgPresu.TextMatrix(N, 6)) < CCur(fgPresu.TextMatrix(N, 9)) Then
'''    If vAcumu < CCur(fgPresu.TextMatrix(n, 9)) Then
'''        If vAcumu < CCur(fgPresu.TextMatrix(n, 9)) Then
'''            MsgBox "Monto a reservar es mayor al presupuesto" & vbCr & "en el periodo de " & fgPresu.TextMatrix(n, 2), vbInformation, "Aviso"
'''            Exit Sub
'''        Else
'''            If MsgBox("Monto a reservar es mayor al presupuesto" & vbCr & "en el periodo de " & fgPresu.TextMatrix(n, 2) & vbCr & "Desea contiuar ya que en el acumulado todavia es inferior ?", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
'''                vAcumu = vAcumu - CCur(fgPresu.TextMatrix(n, 9))
'''            Else
'''                Exit Sub
'''            End If
'''        End If
'''    End If
'''Next
'''If MsgBox("Esta seguro que desea reservar en esta cuenta ?", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
'''    dbCmact.BeginTrans
'''    'Actualiza la cta cnt, presu, tipcambio en LObtenDet
'''    tmpSql = "Update LObtenDet Set cCtaCnt = '" & Trim(txtCtaCnt) & "', cPresu = '" & pPresu & "', nTipCam = " & Val(txtTipCam) & "  Where cObjetoCod = '" & pBSCod & "'And cNroObt = '" & pReqNro & "'"
'''    dbCmact.Execute tmpSql
'''
'''    'Reserva monto en PPresupu
'''    For m = 1 To fgPresu.Rows - 1
'''        If CCur(fgPresu.TextMatrix(m, 9)) > 0 Then
'''            tmpSql = "Update PPresupu Set nMonRes =  nMonRes +  " & CCur(fgPresu.TextMatrix(m, 9)) & " " & _
'''                " Where cAno = '" & pAno & "' And cPresu = '" & pPresu & "' " & _
'''                " And cCodRub = '" & fgPresu.TextMatrix(m, 3) & "' And cPeriodo = '" & fgPresu.TextMatrix(m, 1) & "'" & _
'''                " "
'''            dbCmact.Execute tmpSql
'''        End If
'''    Next
'''    dbCmact.CommitTrans
'''
'''    cboPrePri.Enabled = False
'''    cboCta.Enabled = False
'''    cmdGrabar.Enabled = False
'''    cmdModificar.Enabled = True
'''    pRpta = Trim(txtCtaCnt)
'''
'''    Call CargaPreCta(txtCtaCnt)
'''
'''    fgPresu.Col = 9
'''    For m = 1 To fgPresu.Rows - 1
'''        fgPresu.Row = m
'''        fgPresu.CellBackColor = &H80000018     ' color de los grabados
'''    Next
'''End If
'''End Sub

'''Private Sub CmdModificar_Click()
'''Dim tmpSql As String
'''Dim m As Integer
'''If MsgBox("Esta seguro que desea eliminar la reserva de esta Cuenta ?", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
'''
'''    dbCmact.BeginTrans
'''    'Actualiza la cta cnt, presu, tipcambio en LObtenDet
'''    tmpSql = "Update LObtenDet Set cCtaCnt = Null , cPresu = Null , nTipCam = Null  Where cObjetoCod = '" & pBSCod & "'And cNroObt = '" & pReqNro & "'"
'''    dbCmact.Execute tmpSql
'''
'''    'Resta monto en Reserva de PPresupu
'''    For m = 1 To fgPresu.Rows - 1
'''        If CCur(fgPresu.TextMatrix(m, 9)) > 0 Then
'''            tmpSql = "Update PPresupu Set nMonRes =  nMonRes -  " & CCur(fgPresu.TextMatrix(m, 9)) & " " & _
'''                " Where cAno = '" & pAno & "' And cPresu = '" & pPresu & "' " & _
'''                " And cCodRub = '" & fgPresu.TextMatrix(m, 3) & "' And cPeriodo = '" & fgPresu.TextMatrix(m, 1) & "'" & _
'''                " "
'''            dbCmact.Execute tmpSql
'''        End If
'''    Next
'''    dbCmact.CommitTrans
'''
'''    pRpta = ""
'''    cboPrePri.Enabled = True
'''    cboCta.Enabled = True
'''    cmdGrabar.Enabled = True
'''    txtTipCam = Format(Val(FuncGnral("Select nValFijoDia Campo From tipcambio Where datediff(dd, dfeccamb, '" & Format(gdFecSis, gsformatofecha) & "') = 0 ")), "#0.000")
'''    cmdModificar.Enabled = False
'''
'''    Call CargaPreCta(txtCtaCnt)
'''End If
'''End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub


'''Private Function VeriCta(ByVal cCta As String) As Boolean
'''Dim tmpSql As String
'''Dim vCant As Integer
'''VeriCta = False
''''Se verifica si la cta. cnt. esta la Tabla CtaCont
'''vCant = Val(FunGnral("Select  Count(*) Campo From CtaCont Where cCtaContCod like '" & cCta & "%'"))
'''If vCant = 1 Then
'''    VeriCta = True
'''ElseIf vCant = 0 Then
'''    MsgBox "Cuenta no existe", vbInformation, "Aviso"
'''ElseIf vCant > 1 Then
'''    MsgBox "Cuenta no esta en el Ultimo Nivel", vbInformation, "Aviso"
'''End If
'''End Function

'''Private Sub CargaPreCta(ByVal sCtaCnt As String)
'''Dim tmpReg As New ADODB.Recordset
'''Dim tmpSql As String
'''Dim x As Integer
'''Dim vAcumu As Currency
'''Call MSHFlex(fgPresu, 10, "Item-cCodPer-Periodo-Rubro Pres.-Inicial-Ejecutado-Actual-Act.Acumul.-Reservado-Por Reserv.", "450-0-900-0-1100-1100-1100-1100-1100-1100", "C-L-L-L-R-R-R-R-R-R")
'''tmpSql = " SELECT t.cCodTab, t.cNomTab, b.cCodRub, b.cDesRub, p.nMonIni, p.nMonAct, p.nMonEje, p.nMonRes " & _
'''    " FROM TablaCod T Left Join PPresupu P ON t.ccodtab = p.cPeriodo  " & _
'''    " Inner Join PRubCta R On p.cAno = r.cAno and p.cPresu = r.cPresu and p.cCodRub = r.cCodRub " & _
'''    " AND r.cAno = '" & pAno & "' and  p.cmoneda = '1' inner join PRubro B On r.ccodrub=b.ccodrub" & _
'''    " Where t.cCodTab like 'P3__' AND t.cCodTab <> 'P300' and p.cPresu = '" & Right(Trim(cboPrePri), 4) & "' " & _
'''    " and substring(cctacnt,1,2) = '" & Left(sCtaCnt, 2) & "' and substring(cctacnt,4,len(cctacnt)-3) = '" & Mid(sCtaCnt, 4) & "' Order By b.cdesrub, t.cCodTab "
'''tmpReg.Open tmpSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'''If (tmpReg.BOF Or tmpReg.EOF) Then
'''Else
'''    With tmpReg
'''        txtNomRub = !cDesRub
'''        Do While Not tmpReg.EOF
'''            x = x + 1
'''            vAcumu = vAcumu + Val(!nMonAct) - Val(!nMonRes)
'''            AdicionaRow fgPresu, x
'''            fgPresu.Row = fgPresu.Rows - 1
'''            fgPresu.TextMatrix(x, 0) = x
'''            fgPresu.TextMatrix(x, 1) = !cCodTab
'''            fgPresu.TextMatrix(x, 2) = Trim(!cNomTab)
'''            fgPresu.TextMatrix(x, 3) = Trim(!cCodRub)
'''            fgPresu.TextMatrix(x, 4) = Format(!nMonIni, "#,##0.00")
'''            fgPresu.TextMatrix(x, 5) = Format(!nMonEje, "#,##0.00")
'''            fgPresu.TextMatrix(x, 6) = Format(!nMonAct, "#,##0.00")
'''            fgPresu.TextMatrix(x, 7) = Format(vAcumu, "#,##0.00")
'''            fgPresu.TextMatrix(x, 8) = Format(!nMonRes, "#,##0.00")
'''            If pMone = "1" Then
'''                fgPresu.TextMatrix(x, 9) = Format(FuncGnral("Select (p.nCantApro * d.nMonPre) Campo From LObtenDet D Inner Join LObtenPer P On d.cNroObt = p.cnroobt and d.cobjetocod = p.cobjetocod Where p.cNroObt = '" & pReqNro & "' and p.cObjetocod = '" & pBSCod & "' and p.cPeriodo = '" & !cCodTab & "'"), "#,##0.00")
'''            Else
'''                fgPresu.TextMatrix(x, 9) = Format(FuncGnral("Select (p.nCantApro * d.nMonPre * " & Val(txtTipCam) & ") campo From LObtenDet D Inner Join LObtenPer P On d.cNroObt = p.cnroobt and d.cobjetocod = p.cobjetocod Where p.cNroObt = '" & pReqNro & "' and p.cObjetocod = '" & pBSCod & "' and p.cPeriodo = '" & !cCodTab & "'"), "#,##0.00")
'''            End If
'''            'If CCur(fgPresu.TextMatrix(X, 9)) > 0 Then
'''            '    fgPresu.TextMatrix(X, 8) = Format(CCur(fgPresu.TextMatrix(X, 8)) - CCur(fgPresu.TextMatrix(X, 9)), "#,##0.00")
'''            'End If
'''            .MoveNext
'''        Loop
'''    End With
'''End If
'''tmpReg.Close
'''Set tmpReg = Nothing
'''End Sub
