VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCtaContMantenimiento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Cuentas Contables"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   Icon            =   "frmCtaContMantenimiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6480
      TabIndex        =   11
      Top             =   7095
      Width           =   975
   End
   Begin VB.ComboBox cboProd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   15
      Width           =   4035
   End
   Begin VB.Frame fraAS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Asiento Contable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4095
      Left            =   30
      TabIndex        =   12
      Top             =   2955
      Width           =   7425
      Begin VB.ComboBox cmbConcepto 
         Height          =   315
         Left            =   5145
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3675
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   6360
         TabIndex        =   4
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   6360
         TabIndex        =   5
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox TxtTipOpe 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   6
         TabIndex        =   6
         Top             =   3690
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.TextBox TxtCodCnt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1065
         MaxLength       =   22
         TabIndex        =   7
         Top             =   3690
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.TextBox txtDEBEHABER 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3555
         MaxLength       =   1
         TabIndex        =   8
         Top             =   3690
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.TextBox TxtTipPers 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4365
         MaxLength       =   1
         TabIndex        =   9
         Top             =   3690
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   6360
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   6360
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexAS 
         Height          =   3405
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   6006
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         SelectionMode   =   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).BandIndent=   1
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   6360
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   6360
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexOpe 
      Height          =   2415
      Left            =   30
      TabIndex        =   1
      Top             =   540
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   4260
      _Version        =   393216
      Rows            =   4
      FixedCols       =   0
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).BandIndent=   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Producto"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   645
   End
End
Attribute VB_Name = "frmCtaContMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnFilaAnt As Long
Dim lbNuevo As Boolean
Dim lnRango As Long

'Dim oCon As COMConecta.DCOMConecta

Private Sub FillProducto(ByVal pnFiltro As Integer)
    'Dim oCon As COMDConstSistema.NCOMConstSistema
    'Set oCon = New COMDConstSistema.NCOMConstSistema
    'Dim lnFiltro As Integer
    
    'lnFiltro = oCon.LeeConstSistema(103)
    
    cboProd.AddItem "AHORROS" & Space(100) & "2%"
    cboProd.AddItem "COMPRA VENTA" & Space(100) & "90002[23]%"
    cboProd.AddItem "CREDITOS" & Space(100) & "10%"
    If pnFiltro = 1 Then
        cboProd.AddItem "PIGNORATICIO" & Space(100) & "12%"
    ElseIf pnFiltro = 2 Then
        cboProd.AddItem "PIGNORATICIO" & Space(100) & "15%"
    Else
        cboProd.AddItem "PIGNORATICIO" & Space(100) & "12%"
        cboProd.AddItem "PIGNORATICIO" & Space(100) & "15%"
    End If
    cboProd.AddItem "RECUPERACIONES" & Space(100) & "13%"
    cboProd.AddItem "CARTA FIANZA" & Space(100) & "14%"
    cboProd.AddItem "OTRAS OPERACIONES" & Space(100) & "3%"
    cboProd.AddItem "TODOS" & Space(100) & "%"
    cboProd.ListIndex = 0
End Sub

Private Sub cboProd_Click()
    Dim lsCodProd As String
    lsCodProd = Trim(Right(cboProd.Text, 10))
    CargaOperacion lsCodProd
    GetAsiento Me.FlexOpe.TextMatrix(Me.FlexOpe.Row, 0)
End Sub

Private Sub cmdCancelar_Click()
    Me.FlexAS.Enabled = True
    Me.FlexOpe.Enabled = True
    TxtCodCnt.Visible = False
    txtDEBEHABER.Visible = False
    TxtTipPers.Visible = False
    TxtTipOpe.Visible = False
    cmbConcepto.Visible = False

    CmdNuevo.Visible = True
    CmdEditar.Visible = True
    cmdgrabar.Visible = False
    cmdCancelar.Visible = False

End Sub

Private Sub CmdEditar_Click()
    If Trim(FlexAS.TextMatrix(FlexAS.Row, 0)) = "" Then
        MsgBox "No existe ninguna Cuenta para esta la Operación : " & Trim(FlexOpe.TextMatrix(Me.FlexOpe.Row, 1)) & " (" & Trim(FlexOpe.TextMatrix(Me.FlexOpe.Row, 0)) & ")", vbInformation, "Aviso"
        Exit Sub
    End If
    
    lbNuevo = False
    
    CmdNuevo.Visible = False
    CmdEditar.Visible = False
    cmdgrabar.Visible = True
    cmdCancelar.Visible = True
    
    TxtTipOpe.Text = Trim(FlexAS.TextMatrix(FlexAS.Row, 0))
    TxtCodCnt.Text = Trim(FlexAS.TextMatrix(FlexAS.Row, 1))
    txtDEBEHABER.Text = Trim(FlexAS.TextMatrix(FlexAS.Row, 2))
    Me.TxtTipPers.Text = Trim(FlexAS.TextMatrix(FlexAS.Row, 3))
    UbicaCombo cmbConcepto, Trim(FlexAS.TextMatrix(FlexAS.Row, 4))

    Me.FlexAS.Enabled = False
    Me.FlexOpe.Enabled = False
    TxtCodCnt.Visible = True
    txtDEBEHABER.Visible = True
    TxtTipPers.Visible = True
    cmbConcepto.Visible = True
    TxtTipOpe.Visible = True

End Sub

Private Sub cmdeliminar_Click()
    'Dim sqlE As String
    Dim oCta As COMDContabilidad.DCOMCtaCont
    
    If Trim(FlexAS.TextMatrix(FlexAS.Row, 0)) = "" Then
        MsgBox "No existe ninguna Cuenta para la Operación : " & Trim(FlexOpe.TextMatrix(Me.FlexOpe.Row, 1)) & " (" & Trim(FlexOpe.TextMatrix(Me.FlexOpe.Row, 0)) & ")", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("Desea Realizar los Cambios?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Set oCta = New COMDContabilidad.DCOMCtaCont
        Call oCta.EliminaCtaContMant(Trim(FlexAS.TextMatrix(FlexAS.Row, 0)), Trim(FlexAS.TextMatrix(FlexAS.Row, 1)), _
                                    Trim(FlexAS.TextMatrix(FlexAS.Row, 2)), Trim(FlexAS.TextMatrix(FlexAS.Row, 3)), _
                                    Trim(FlexAS.TextMatrix(FlexAS.Row, 4)))
        'sqlE = "Delete OpeCtaNeg " _
             & " Where cOpeCod =  '" & Trim(FlexAS.TextMatrix(FlexAS.Row, 0)) & "'" _
             & " and cCtaContCod =  '" & Trim(FlexAS.TextMatrix(FlexAS.Row, 1)) & "'" _
             & " and cOpeCtaDH =  '" & Trim(FlexAS.TextMatrix(FlexAS.Row, 2)) & "'" _
             & " and nPersoneria =  '" & Trim(FlexAS.TextMatrix(FlexAS.Row, 3)) & "'" _
             & " And nConcepto =  '" & Trim(FlexAS.TextMatrix(FlexAS.Row, 4)) & "'"
        
        'oCon.Ejecutar sqlE
        
        GetAsiento Me.FlexOpe.TextMatrix(Me.FlexOpe.Row, 0)
    End If
End Sub

Private Sub cmdGrabar_Click()
'Dim sqlG As String
'Dim oMov As COMDMov.DCOMMov
'Set oMov = New COMDMov.DCOMMov
Dim oCtaCont As COMDContabilidad.DCOMCtaCont
     
If Me.TxtCodCnt = "" Or Me.txtDEBEHABER = "" Or Me.TxtTipOpe = "" Or Me.TxtTipPers = "" Then
    MsgBox "No puede tener ningún Campo en blanco al realizar la Grabación"
    If Me.TxtCodCnt.Text = "" Then TxtCodCnt.SetFocus
    If Me.txtDEBEHABER.Text = "" Then TxtCodCnt.SetFocus
    If Me.TxtTipOpe.Text = "" Then TxtCodCnt.SetFocus
    If Me.TxtTipPers.Text = "" Then TxtCodCnt.SetFocus
    If Me.cmbConcepto.Text = "" Then cmbConcepto.SetFocus
End If
    
    If MsgBox("Desea Realizar los Cambios?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Set oCtaCont = New COMDContabilidad.DCOMCtaCont
        
        If lbNuevo Then
            'sqlG = "Insert OpeCtaNeg (cOpeCod,nConcepto,nPersoneria,cCtaContCod,cOpeCtaDH,cUltimaActualizacion) " _
                 & "Values('" & TxtTipOpe & "','" & Trim(Right(Me.cmbConcepto.Text, 6)) & "','" & TxtTipPers.Text & "','" & TxtCodCnt.Text & "','" & txtDEBEHABER & "','" & oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser) & "')"
            'oCon.Ejecutar sqlG
            Call oCtaCont.InsertaCtaContMant(TxtTipOpe.Text, Trim(Right(Me.cmbConcepto.Text, 6)), TxtTipPers.Text, TxtCodCnt.Text, txtDEBEHABER.Text, gdFecSis, gsCodAge, gsCodUser)
        Else
            'sqlG = " Update OpeCtaNeg " _
                 & " set cOpeCod = '" & TxtTipOpe.Text & "'," _
                 & " cCtaContCod = '" & TxtCodCnt.Text & "'," _
                 & " cOpeCtaDH = '" & txtDEBEHABER.Text & "'," _
                 & " nPersoneria = '" & TxtTipPers.Text & "'," _
                 & " nConcepto = '" & Trim(Right(Me.cmbConcepto.Text, 6)) & "'," _
                 & " cUltimaActualizacion = '" & oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser) & "'" _
                 & " Where cOpeCod =  '" & Trim(FlexAS.TextMatrix(FlexAS.Row, 0)) & "'" _
                 & " and cCtaContCod =  '" & Trim(FlexAS.TextMatrix(FlexAS.Row, 1)) & "'" _
                 & " and cOpeCtaDH =  '" & Trim(FlexAS.TextMatrix(FlexAS.Row, 2)) & "'" _
                 & " and nPersoneria =  '" & Trim(FlexAS.TextMatrix(FlexAS.Row, 3)) & "'" _
                 & " and nConcepto =  '" & Trim(FlexAS.TextMatrix(FlexAS.Row, 4)) & "'"
            
            Call oCtaCont.ModificaCtaContMant(TxtTipOpe.Text, Trim(Right(Me.cmbConcepto.Text, 6)), _
                                        TxtTipPers.Text, TxtCodCnt.Text, txtDEBEHABER.Text, gdFecSis, gsCodAge, gsCodUser, _
                                        Trim(FlexAS.TextMatrix(FlexAS.Row, 0)), Trim(FlexAS.TextMatrix(FlexAS.Row, 4)), _
                                        Trim(FlexAS.TextMatrix(FlexAS.Row, 3)), Trim(FlexAS.TextMatrix(FlexAS.Row, 1)), _
                                        Trim(FlexAS.TextMatrix(FlexAS.Row, 2)))
            'oCon.Ejecutar sqlG
        End If
        Set oCtaCont = Nothing
        
        GetAsiento Me.FlexOpe.TextMatrix(Me.FlexOpe.Row, 0)
        
        CmdNuevo.Visible = True
        CmdEditar.Visible = True
        cmdgrabar.Visible = False
        cmdCancelar.Visible = False
     
     End If

    cmdCancelar_Click
End Sub

Private Sub cmdImprimir_Click()
    Dim lsCodProd As String
    Dim oPrevio As previo.clsprevio
    Dim lsCadena As String
    Dim oImpre As COMNContabilidad.NCOMContImprimir
    
    
    lsCodProd = Trim(Right(cboProd.Text, 10))
    
    Set oImpre = New COMNContabilidad.NCOMContImprimir
    lsCadena = oImpre.ImpreAsientoCtaContMant("Listado de Asientos de Operaciones", 0, 0, lsCodProd, gsNomAge, gsNomCmac, gdFecSis, lnRango, gImpresora)
    Set oImpre = Nothing
    
    Set oPrevio = New previo.clsprevio
    oPrevio.Show lsCadena, "Listado de Asientos de Operaciones", True, 66
    Set oPrevio = Nothing
End Sub

Private Sub cmdNuevo_Click()
    lbNuevo = True
    
    CmdNuevo.Visible = False
    CmdEditar.Visible = False
    cmdgrabar.Visible = True
    cmdCancelar.Visible = True
    
    FlexOpe.Enabled = False
    FlexAS.Enabled = False
    TxtTipOpe = FlexOpe.TextMatrix(FlexOpe.Row, 0)
    TxtCodCnt.Visible = True
    txtDEBEHABER.Visible = True
    TxtTipPers.Visible = True
    TxtTipOpe.Visible = True
    cmbConcepto.Visible = True

    TxtCodCnt = ""
    txtDEBEHABER = ""
    TxtTipPers = ""
    
    TxtCodCnt.SetFocus
    
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub FlexOpe_LeaveCell()
    lnFilaAnt = FlexOpe.Row
End Sub

Private Sub FlexOpe_RowColChange()
    If FlexOpe.Row <> lnFilaAnt Then
        GetAsiento Me.FlexOpe.TextMatrix(Me.FlexOpe.Row, 0)
    End If
End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    'Dim sql As String
    'Set rs = New ADODB.Recordset
    'Set oCon = New COMConecta.DCOMConecta
    Dim nFiltro As Integer
    Dim oCta As COMDContabilidad.DCOMCtaCont
    
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    
    'oCon.AbreConexion
    
    'sql = "Select convert(varchar(8), nPrdConceptoCod) + ' - ' + cDescripcion, nPrdConceptoCod from ProductoConcepto Order By nPrdConceptoCod"
    'Set rs = oCon.CargaRecordSet(sql)
    Set oCta = New COMDContabilidad.DCOMCtaCont
    Call oCta.CargarControlesCtaContMant(rs, nFiltro)
    Set oCta = Nothing
    CambiaTamañoCombo Me.cmbConcepto, 300
    CargaCombo Me.cmbConcepto, rs
    
    Call FillProducto(nFiltro)
    Ini
    lnRango = 57
End Sub

Private Sub CargaOperacion(psCodPro As String)
    'Dim sqlOpe As String
    Dim rsOpe As New ADODB.Recordset
    'Set rsOpe = New ADODB.Recordset
    Dim oCta As COMDContabilidad.DCOMCtaCont
    
    'sqlOpe = "Select cOpeCod, cOpeDesc From OpeTpo Where cOpeCod Like '" & psCodPro & "' and RTrim(cOpeCod) like '______' Order by cOpeCod"
    'Set rsOpe = oCon.CargaRecordSet(sqlOpe)
    
    Set oCta = New COMDContabilidad.DCOMCtaCont
    Set rsOpe = oCta.CargaOperacion(psCodPro)
    Set oCta = Nothing
    If rsOpe.EOF Then
        MsgBox "La Tabla Operación esta Vacia para el producto " & psCodPro, vbInformation, "Aviso"
    Else
        Me.FlexOpe.Rows = 2
        While Not rsOpe.EOF
            FlexOpe.TextMatrix(FlexOpe.Rows - 1, 0) = rsOpe!cOpecod
            FlexOpe.TextMatrix(FlexOpe.Rows - 1, 1) = rsOpe!cOpedesc
            Me.FlexOpe.Rows = Me.FlexOpe.Rows + 1
            rsOpe.MoveNext
        Wend
        Me.FlexOpe.Rows = Me.FlexOpe.Rows - 1
        rsOpe.Close
    End If
    
    Set rsOpe = Nothing
End Sub

Private Sub Ini()
    FlexOpe.TextMatrix(0, 0) = "Cod.Ope."
    Me.FlexOpe.ColWidth(0) = 230 * 6
    FlexOpe.TextMatrix(0, 1) = "Nombre Operación"
    Me.FlexOpe.ColWidth(1) = 230 * 17
    
    FlexAS.TextMatrix(0, 0) = "Cod.Ope."
    Me.FlexAS.ColWidth(0) = 230 * 5
    FlexAS.TextMatrix(0, 1) = "Codigo Contable"
    Me.FlexAS.ColWidth(1) = 230 * 10
    FlexAS.TextMatrix(0, 2) = "Debe/Haber"
    Me.FlexAS.ColWidth(2) = 230 * 4
    FlexAS.TextMatrix(0, 3) = "Tip.Per"
    Me.FlexAS.ColWidth(3) = 230 * 3
    FlexAS.TextMatrix(0, 4) = "Concepto"
    Me.FlexAS.ColWidth(4) = 230 * 3

End Sub


Private Sub GetAsiento(psCodOpe As String)
    'Dim sqlAS As String
    Dim rsAS As New ADODB.Recordset
    Dim oCta As COMDContabilidad.DCOMCtaCont
    'sqlAS = "Select cOpeCod, cCtaContCod, cOpeCtaDH, nConcepto, nPersoneria  from OpeCtaNeg where cOpeCod = '" & psCodOpe & "' Order By nPersoneria, nConcepto, cOpeCtaDH  "
    'Set rsAS = oCon.CargaRecordSet(sqlAS)
    
    FlexAS.Rows = 1
    Me.FlexAS.Rows = 2
    Me.FlexAS.FixedRows = 1
    
    Set oCta = New COMDContabilidad.DCOMCtaCont
    Set rsAS = oCta.GetAsiento(psCodOpe)
    Set oCta = Nothing
    If Not rsAS.EOF Then
    
        While Not rsAS.EOF
            Me.FlexAS.TextMatrix(Me.FlexAS.Rows - 1, 0) = Trim(rsAS!cOpecod)
            Me.FlexAS.TextMatrix(Me.FlexAS.Rows - 1, 1) = Trim(rsAS!cCtaContCod)
            Me.FlexAS.TextMatrix(Me.FlexAS.Rows - 1, 2) = Trim(rsAS!cOpeCtaDH)
            Me.FlexAS.TextMatrix(Me.FlexAS.Rows - 1, 3) = Trim(rsAS!nPersoneria)
            Me.FlexAS.TextMatrix(Me.FlexAS.Rows - 1, 4) = Trim(rsAS!nConcepto)
            Me.FlexAS.Rows = Me.FlexAS.Rows + 1
            rsAS.MoveNext
        Wend
        Me.FlexAS.Rows = Me.FlexAS.Rows - 1
        rsAS.Close
    End If
    Set rsAS = Nothing
End Sub

Private Sub TxtCodCnt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtDEBEHABER.SetFocus
    End If
    'KeyAscii = Asc(UCase(Chr(ValdiaText(KeyAscii))))
End Sub

Private Sub txtDEBEHABER_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.TxtTipPers.SetFocus
        Exit Sub
    End If
    If KeyAscii = 8 Then Exit Sub
    If UCase(Chr(KeyAscii)) = "H" Or UCase(Chr(KeyAscii)) = "D" Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Else
        KeyAscii = 0
    End If
    
End Sub

Private Sub TxtTipPers_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then Me.cmbConcepto.SetFocus
    If InStr(1, "0123456789", Chr(KeyAscii), vbTextCompare) = 0 Then KeyAscii = 0
End Sub

'Private Function ImpreAsiento(psTitulo As String, pnPagina As Long, pnItem As Long, psCodPro As String) As String
'    Dim lsCad As String
'    Dim sqlOpe As String
'    Dim rsOpe As New ADODB.Recordset
'    Dim lbBan As Boolean
'    Dim lsD As String
'    Dim lsH As String
'    Dim rsD As New ADODB.Recordset
'    Dim rsH As New ADODB.Recordset
'    Dim lsNomOpe As String * 55
'    Dim sqlOC As String
'    Dim rsOC As New ADODB.Recordset
'
'    sqlOpe = "Select cOpeCod, cOpeDesc From OpeTpo where cOpeCod like '" & psCodPro & "' And RTrim(cOpeCod) like '______' order by cOpeCod"
'    Set rsOpe = oCon.CargaRecordSet(sqlOpe)
'
'    If Not RSVacio(rsOpe) Then
'        lsCad = ""
'        lsCad = lsCad & CabeceraPagina(psTitulo, pnPagina, pnItem, gsNomAge, gsNomCmac, gdFecSis, gMonedaNacional, False)
'        lsCad = lsCad & Encabezado("Cuenta;10;Descripcion;25; ;5;Debe;20;Haber;20;Concepto;25;Tip.Pers;10;Des_Concepto;15; ;1;", pnItem)
'
'        While Not rsOpe.EOF
'            lbBan = True
'
'            lsD = ""
'            lsH = ""
'
'            sqlOC = " Select distinct OPN.nPersoneria, OPN.nConcepto, PC.cDescripcion From OpeCtaNeg OPN " _
'                  & " Inner Join ProductoConcepto PC On PC.nPrdConceptoCod = OPN.nConcepto  " _
'                  & " Where OPN.cOpeCod = '" & Trim(rsOpe!cOpecod) & " ' "
'            Set rsOC = oCon.CargaRecordSet(sqlOC)
'
'            lsNomOpe = rsOpe!cOpedesc
'
'            If RSVacio(rsOC) Then
'                pnItem = pnItem + 1
'                lsCad = lsCad & Space(2) & Trim(rsOpe!cOpecod) _
'                              & Space(1) & ImpreFormat(Trim(Left(lsNomOpe, 40)), 40) _
'                              & ImpreFormat(lsD, 22) _
'                              & ImpreFormat(lsH, 22) _
'                              & oImpresora.gPrnSaltoLinea
'
'                If pnItem = lnRango Then
'                    lsCad = lsCad & oImpresora.gPrnSaltoPagina
'                    lsCad = lsCad & CabeceraPagina(psTitulo, pnPagina, pnItem, gsNomAge, gsNomCmac, gdFecSis, gMonedaNacional, False)
'                    lsCad = lsCad & Encabezado("Cuenta;10;Descripcion;25; ;5;Debe;20;Haber;20;Concepto;25;Tip.Pers;10;Des_Concepto;15; ;1;", pnItem)
'                End If
'            Else
'                While Not rsOC.EOF
'                    Set rsD = GetOpeCuenta(Trim(rsOpe!cOpecod), "D", rsOC!nPersoneria, rsOC!nConcepto)
'                    Set rsH = GetOpeCuenta(Trim(rsOpe!cOpecod), "H", rsOC!nPersoneria, rsOC!nConcepto)
'
'                    While Not rsD.EOF And Not rsH.EOF
'                        lsD = Trim(rsD.Fields(0))
'                        lsH = Trim(rsH.Fields(0))
'
'                        If lbBan Then
'                            pnItem = pnItem + 1
'                            lsCad = lsCad & Space(2) & Trim(rsOpe!cOpecod) _
'                                          & Space(1) & ImpreFormat(Trim(Left(lsNomOpe, 40)), 40) _
'                                          & ImpreFormat(lsD, 22) _
'                                          & ImpreFormat(lsH, 22) _
'                                          & ImpreFormat(Trim(rsOC!nConcepto), 8) & Space(1) & Trim(rsOC!nPersoneria) & Space(2) & ImpreFormat(Trim(rsOC!cDescripcion), 20) & oImpresora.gPrnSaltoLinea
'
'                            lbBan = False
'                            If pnItem = lnRango Then
'                                lsCad = lsCad & oImpresora.gPrnSaltoPagina
'                                lsCad = lsCad & CabeceraPagina(psTitulo, pnPagina, pnItem, gsNomAge, gsNomCmac, gdFecSis, gMonedaNacional, False)
'                                lsCad = lsCad & Encabezado("Cuenta;10;Descripcion;25; ;5;Debe;20;Haber;20;Concepto;25;Tip.Pers;10;Des_Concepto;15; ;1;", pnItem)
'                            End If
'                        Else
'                            pnItem = pnItem + 1
'                            lsCad = lsCad & Space(51) _
' _
'                                          & ImpreFormat(lsD, 22) _
'                                          & ImpreFormat(lsH, 22) _
'                                          & ImpreFormat(Trim(rsOC!nConcepto), 8) & Space(1) & Trim(rsOC!nPersoneria) & Space(2) & ImpreFormat(Trim(rsOC!cDescripcion), 20) & oImpresora.gPrnSaltoLinea
'
'                            If pnItem = lnRango Then
'                                lsCad = lsCad & oImpresora.gPrnSaltoPagina
'                                lsCad = lsCad & CabeceraPagina(psTitulo, pnPagina, pnItem, gsNomAge, gsNomCmac, gdFecSis, gMonedaNacional, False)
'                                lsCad = lsCad & Encabezado("Cuenta;10;Descripcion;25; ;5;Debe;20;Haber;20;Concepto;25;Tip.Pers;10;Des_Concepto;15; ;1;", pnItem)
'                            End If
'                        End If
'                        rsD.MoveNext
'                        rsH.MoveNext
'                    Wend
'
'                    While Not rsD.EOF
'                        lsD = Trim(rsD.Fields(0))
'                        lsH = ""
'
'                        If lbBan Then
'                            pnItem = pnItem + 1
'                            lsCad = lsCad & Space(2) & Trim(rsOpe!cOpecod) _
'                                          & Space(1) & ImpreFormat(Trim(Left(lsNomOpe, 40)), 40) _
'                                          & ImpreFormat(lsD, 22) _
'                                          & ImpreFormat(lsH, 22) _
'                                          & ImpreFormat(Trim(rsOC!nConcepto), 8) & Space(1) & Trim(rsOC!nPersoneria) & Space(2) & ImpreFormat(Trim(rsOC!cDescripcion), 20) & oImpresora.gPrnSaltoLinea
'
'                            lbBan = False
'                            If pnItem = lnRango Then
'                                lsCad = lsCad & oImpresora.gPrnSaltoPagina
'                                lsCad = lsCad & CabeceraPagina(psTitulo, pnPagina, pnItem, gsNomAge, gsNomCmac, gdFecSis, gMonedaNacional, False)
'                                lsCad = lsCad & Encabezado("Cuenta;10;Descripcion;25; ;5;Debe;20;Haber;20;Concepto;25;Tip.Pers;10;Des_Concepto;15; ;1;", pnItem)
'                            End If
'                        Else
'                            pnItem = pnItem + 1
'                            lsCad = lsCad & Space(51) _
' _
'                                          & ImpreFormat(lsD, 22) _
'                                          & ImpreFormat(lsH, 22) _
'                                          & ImpreFormat(Trim(rsOC!nConcepto), 8) & Space(1) & Trim(rsOC!nPersoneria) & Space(2) & ImpreFormat(Trim(rsOC!cDescripcion), 20) & oImpresora.gPrnSaltoLinea
'
'                            If pnItem = lnRango Then
'                                lsCad = lsCad & oImpresora.gPrnSaltoPagina
'                                lsCad = lsCad & CabeceraPagina(psTitulo, pnPagina, pnItem, gsNomAge, gsNomCmac, gdFecSis, gMonedaNacional, False)
'                                lsCad = lsCad & Encabezado("Cuenta;10;Descripcion;25; ;5;Debe;20;Haber;20;Concepto;25;Tip.Pers;10;Des_Concepto;15; ;1;", pnItem)
'                            End If
'                        End If
'                        rsD.MoveNext
'                    Wend
'
'                    While Not rsH.EOF
'                        lsD = ""
'                        lsH = Trim(rsH.Fields(0))
'
'                        If lbBan Then
'                            pnItem = pnItem + 1
'                            lsCad = lsCad & Space(2) & Trim(rsOpe!cOpecod) _
'                                          & Space(1) & ImpreFormat(Trim(Left(lsNomOpe, 40)), 40) _
'                                          & ImpreFormat(lsD, 22) _
'                                          & ImpreFormat(lsH, 22) _
'                                          & ImpreFormat(Trim(rsOC!nConcepto), 8) & Space(1) & Trim(rsOC!nPersoneria) & Space(2) & ImpreFormat(Trim(rsOC!cDescripcion), 20) & oImpresora.gPrnSaltoLinea
'
'                            lbBan = False
'                            If pnItem = lnRango Then
'                                lsCad = lsCad & oImpresora.gPrnSaltoPagina
'                                lsCad = lsCad & CabeceraPagina(psTitulo, pnPagina, pnItem, gsNomAge, gsNomCmac, gdFecSis, gMonedaNacional, False)
'                                lsCad = lsCad & Encabezado("Cuenta;10;Descripcion;25; ;5;Debe;20;Haber;20;Concepto;25;Tip.Pers;10;Des_Concepto;15; ;1;", pnItem)
'                            End If
'                        Else
'                            pnItem = pnItem + 1
'                            lsCad = lsCad & Space(51) _
' _
'                                          & ImpreFormat(lsD, 22) _
'                                          & ImpreFormat(lsH, 22) _
'                                          & ImpreFormat(Trim(rsOC!nConcepto), 8) & Space(1) & Trim(rsOC!nPersoneria) & Space(2) & ImpreFormat(Trim(rsOC!cDescripcion), 20) & oImpresora.gPrnSaltoLinea
'
'                            If pnItem = lnRango Then
'                                lsCad = lsCad & oImpresora.gPrnSaltoPagina
'                                lsCad = lsCad & CabeceraPagina(psTitulo, pnPagina, pnItem, gsNomAge, gsNomCmac, gdFecSis, gMonedaNacional, False)
'                                lsCad = lsCad & Encabezado("Cuenta;10;Descripcion;25; ;5;Debe;20;Haber;20;Concepto;25;Tip.Pers;10;Des_Concepto;15; ;1;", pnItem)
'                            End If
'                        End If
'                        rsH.MoveNext
'                    Wend
'
'                    rsOC.MoveNext
'                Wend
'            End If
'
'            rsOC.Close
'            Set rsOC = Nothing
'            DoEvents
'            rsOpe.MoveNext
'        Wend
'
'    End If
'    ImpreAsiento = lsCad
'End Function

'Private Function GetOpeCuenta(psCodOpe As String, psDH As String, psTipPers As String, psConcepto As String) As ADODB.Recordset
'    Dim sqlOC As String
'    Dim rsOC As ADODB.Recordset
'    Set rsOC = New ADODB.Recordset
'
'    sqlOC = "Select cCtaContCod From OpeCtaNeg Where cOpeCod = '" & psCodOpe & "' and cOpeCtaDH = '" & psDH & "' and nPersoneria = " & psTipPers & " And nConcepto = " & psConcepto & ""
'    Set rsOC = oCon.CargaRecordSet(sqlOC)
'
'    Set GetOpeCuenta = rsOC
'    If Not RSVacio(rsOC) Then
'
'        'rsOC.Close
'    End If
'    'Set rsOC = Nothing
'End Function


