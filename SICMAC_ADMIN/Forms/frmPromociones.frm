VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPromociones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes de Promociones"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   Icon            =   "frmPromociones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLPromocionesNOPF 
      Caption         =   "Prom sin PF"
      Height          =   345
      Left            =   4215
      TabIndex        =   14
      Top             =   3540
      Width           =   1110
   End
   Begin VB.TextBox txtparametro 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4905
      TabIndex        =   10
      Text            =   "3"
      Top             =   1275
      Width           =   570
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   4215
      TabIndex        =   5
      Top             =   3930
      Width           =   1110
   End
   Begin VB.CommandButton cmdApetruras 
      Caption         =   "L. Aperturas"
      Height          =   345
      Left            =   4215
      TabIndex        =   4
      Top             =   3150
      Width           =   1110
   End
   Begin VB.CommandButton cmdLPromociones 
      Caption         =   "Promocion"
      Height          =   345
      Left            =   4215
      TabIndex        =   3
      Top             =   2745
      Width           =   1110
   End
   Begin VB.Frame fraFechas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Fechas"
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
      Height          =   1110
      Left            =   3885
      TabIndex        =   7
      Top             =   75
      Width           =   1605
      Begin MSMask.MaskEdBox mskFI 
         Height          =   300
         Left            =   435
         TabIndex        =   1
         Top             =   300
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFF 
         Height          =   300
         Left            =   435
         TabIndex        =   2
         Top             =   675
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblA 
         Caption         =   "A :"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   420
      End
      Begin VB.Label lblDe 
         Caption         =   "De :"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   345
         Width           =   420
      End
   End
   Begin VB.Frame fraAge 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Agencia"
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
      Height          =   4230
      Left            =   60
      TabIndex        =   6
      Top             =   75
      Width           =   3750
      Begin VB.ListBox lstAge 
         Appearance      =   0  'Flat
         Height          =   3855
         Left            =   90
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   285
         Width           =   3555
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flex 
      Height          =   180
      Left            =   4725
      TabIndex        =   12
      Top             =   2700
      Visible         =   0   'False
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   318
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.OLE OleExcel 
      Appearance      =   0  'Flat
      AutoActivate    =   3  'Automatic
      Enabled         =   0   'False
      Height          =   240
      Left            =   4050
      SizeMode        =   1  'Stretch
      TabIndex        =   13
      Top             =   2160
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblDiaG 
      Alignment       =   1  'Right Justify
      Caption         =   "Dias Dem : "
      Height          =   195
      Left            =   3930
      TabIndex        =   11
      Top             =   1320
      Width           =   870
   End
End
Attribute VB_Name = "frmPromociones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Private Sub cmdApetruras_Click()
    Dim Sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean
    Dim lnI As Integer
    Dim lsFiltro As String
    Dim lsCadAge As String
    Dim lsCadConsol As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    On Error GoTo ERROR
    
    If Not Valida(False) Then Exit Sub
    
    oCon.AbreConexion
    
    For lnI = 0 To Me.lstAge.ListCount - 1
        If lstAge.Selected(lnI) Then
            Sql = " Select  PR.cPromocNro, PE.cPersCod, PE.cPersNombre , PR.cCodPromotor,dFecPromo, PR.cComenta, PC.cCtaCod,  CA.dApertura , CA.nSaldoDisp nSaldDisp, PRD.nSaldo nSaldCnt, nPrdEstado, nPlazo  " _
                & " From  PersPromocion PR " _
                & " Inner Join Persona PE On PE.cPersCod = PR.cPersCod " _
                & " Inner Join ProductoPersona PC On PR.cPersCod = PC.cPersCod And Substring(PC.cCtaCod,3,3) In ('232','233','234')" _
                & " Inner Join Captaciones As CA On CA.cCtaCod = PC.cCtaCod " _
                & " Inner Join Producto As PRD On CA.cCtaCod = PRD.cCtaCod" _
                & " Left  Join CaptacPlazoFijo As CPF On CA.cCtaCod = CPF.cCtaCod " _
                & " Where DateAdd(Day,-" & Me.txtParametro.Text & ",PR.dFecPromo) <= CA.dApertura And PR.dFecPromo Between '" & Format(Me.mskFI.Text, gsFormatoFecha) & "' And '" & Format(Me.mskFF.Text, gsFormatoFecha) & "' And PRD.cCtaCod Like '___" & Right(lstAge.List(lnI), 2) & "%'" _
                & " Order By PE.cPersNombre"
            
            If rsTemp.State = 1 Then rsTemp.Close
            
            Set rsTemp = oCon.CargaRecordSet(Sql)
            
            If Not (rsTemp.EOF And rsTemp.BOF) Then
               AdicionaRecordSet rs, rsTemp
            End If
        End If
    Next lnI
    
    
    Me.Flex.Cols = 12
    Me.Flex.Rows = 1
    
    Me.Flex.TextMatrix(0, 0) = "Cod.Prom"
    Me.Flex.TextMatrix(0, 1) = "Cod.Cliente"
    Me.Flex.TextMatrix(0, 2) = "Nom.Cliente"
    Me.Flex.TextMatrix(0, 3) = "Usu.Promotor"
    Me.Flex.TextMatrix(0, 4) = "F.Promocion"
    Me.Flex.TextMatrix(0, 5) = "Comentario"
    Me.Flex.TextMatrix(0, 6) = "Cuenta"
    Me.Flex.TextMatrix(0, 7) = "Apertura"
    Me.Flex.TextMatrix(0, 8) = "Saldo.Disp"
    Me.Flex.TextMatrix(0, 9) = "Saldo.Cont"
    Me.Flex.TextMatrix(0, 10) = "Estado"
    Me.Flex.TextMatrix(0, 11) = "Plazo"
    
    If rs.State = 0 Then
        MsgBox "No se encontraron datos.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    rs.MoveFirst
    If Not RSVacio(rs) Then
        While Not rs.EOF
            Me.Flex.Rows = Me.Flex.Rows + 1
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 0) = "'" & rs.Fields(0)
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 1) = "'" & rs.Fields(1)
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 2) = rs.Fields(2)
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 3) = rs.Fields(3)
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 4) = Format(rs.Fields(4), ": dd/mm/yyyy")
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 5) = rs.Fields(5)
            
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 6) = ": " & rs.Fields(6)
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 7) = Format(rs.Fields(7), ": dd/mm/yyyy")
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 8) = rs.Fields(8)
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 9) = rs.Fields(9)
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 10) = rs.Fields(10)
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 11) = rs.Fields(11)
            
            rs.MoveNext
        Wend
    End If
    
    lsArchivoN = App.path & "\Spooler\" & Format(CDate(Me.mskFF.Text & " " & Time), gsFormatoMovFechaHora) & ".xls"
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
       If lbLibroOpen Then
          Set xlHoja1 = xlLibro.Worksheets(1)
          ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
          
          Call GeneraReporte
    
          OleExcel.Class = "ExcelWorkSheet"
          ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
          OleExcel.SourceDoc = lsArchivoN
          OleExcel.Verb = 1
          OleExcel.Action = 1
          OleExcel.DoVerb -1
       End If
       MousePointer = 0
        
    rs.Close
    Set rs = Nothing
    Exit Sub
ERROR:
    MsgBox Err.Description, vbInformation, "Aviso"

End Sub

Private Sub cmdLPromociones_Click()
    Dim Sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean
    Dim lnI As Integer
    Dim lsFiltro As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    lsFiltro = ""
    For lnI = 0 To Me.lstAge.ListCount - 1
        If lstAge.Selected(lnI) Then
            If lsFiltro = "" Then
                lsFiltro = Right(lstAge.List(lnI), 2)
            Else
                lsFiltro = lsFiltro & "','" & Right(lstAge.List(lnI), 2)
            End If
        End If
    Next lnI
    
    If Not Valida(False) Then Exit Sub
    
    Sql = "  Select PR.cPromocNro, PE.cPersCod, PE.cPersNombre , PR.cCodPromotor,dFecPromo, PR.cComenta, nCapMonSoles, nCapMonDolares, nColMonSoles, nColMonDolares, RTrim(ISNULL(cUbiGeoDescripcion,'')) + ' ' + PE.cPersDireccDomicilio Dir,   'DNI' + ' ' + IsNull(PEI.cPersIDnro,'') Doc, PE.cPersTelefono cTelPers  From PersPromocion PR " _
        & " Inner Join Persona PE On PE.cPersCod = PR.cPersCod " _
        & " Left Join PersID PEI On PE.cPersCod = PEI.cPersCod And cPersIDTpo = 1" _
        & " Left Join UbicacionGeografica ZO ON ZO.cUbiGeoCod = PE.cPersDireccUbiGeo" _
        & " Where PR.dFecPromo  Between '" & Format(Me.mskFI.Text, gsFormatoFecha) & "' And '" & Format(Me.mskFF.Text, gsFormatoFecha) & "' And PR.cCodAge In ('" & lsFiltro & "')  Order By PE.cPersNombre"
        Set rs = oCon.CargaRecordSet(Sql)
        
    Me.Flex.Cols = 13
    Me.Flex.Rows = 1
    
    Me.Flex.TextMatrix(0, 0) = "Cod.Prom"
    Me.Flex.TextMatrix(0, 1) = "Cod.Cliente"
    Me.Flex.TextMatrix(0, 2) = "Nom.Cliente"
    Me.Flex.TextMatrix(0, 3) = "Usu.Promotor"
    Me.Flex.TextMatrix(0, 4) = "F.Promocion"
    Me.Flex.TextMatrix(0, 5) = "Comentario"
    Me.Flex.TextMatrix(0, 6) = "Monto S/. Capta"
    Me.Flex.TextMatrix(0, 7) = "Monto US$/. Capta"
    Me.Flex.TextMatrix(0, 8) = "Monto S/. Coloc"
    Me.Flex.TextMatrix(0, 9) = "Monto US$/. Coloc"
    Me.Flex.TextMatrix(0, 10) = "Direccion"
    Me.Flex.TextMatrix(0, 11) = "DNI"
    Me.Flex.TextMatrix(0, 12) = "Telefono"
    
    If Not RSVacio(rs) Then
        While Not rs.EOF
            Me.Flex.Rows = Me.Flex.Rows + 1
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 0) = "'" & rs.Fields(0)
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 1) = "'" & rs.Fields(1)
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 2) = rs.Fields(2)
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 3) = rs.Fields(3)
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 4) = Format(rs.Fields(4), ": dd/mm/yyyy")
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 5) = rs.Fields(5)
            
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 6) = Format(rs.Fields(6), "#,##0.00")
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 7) = Format(rs.Fields(7), "#,##0.00")
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 8) = Format(rs.Fields(8), "#,##0.00")
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 9) = Format(rs.Fields(9), "#,##0.00")
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 10) = rs.Fields(10)
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 11) = rs.Fields(11)
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 12) = "" & rs.Fields(12)
            
            rs.MoveNext
        Wend
    End If
    
    lsArchivoN = App.path & "\Spooler\" & Format(CDate(Me.mskFF.Text & " " & Time), gsFormatoMovFechaHora) & ".xls"
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
       If lbLibroOpen Then
          Set xlHoja1 = xlLibro.Worksheets(1)
          ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
          
          Call GeneraReporte
    
          OleExcel.Class = "ExcelWorkSheet"
          ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
          OleExcel.SourceDoc = lsArchivoN
          OleExcel.Verb = 1
          OleExcel.Action = 1
          OleExcel.DoVerb -1
       End If
       MousePointer = 0
        
    rs.Close
    Set rs = Nothing
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    Sql = " Select cAgeCod, cAgeDescripcion  From Agencias "
    Set rs = oCon.CargaRecordSet(Sql)
    
    Me.lstAge.Clear
    
    While Not rs.EOF
        Me.lstAge.AddItem rs!cAgeDescripcion & Space(50) & rs!cAgeCod
        rs.MoveNext
    Wend
    
    rs.Close
End Sub

Private Sub Form_Terminate()
    CierraConexion
End Sub

Private Sub mskFI_GotFocus()
    mskFI.SelStart = 0
    mskFI.SelLength = 50
End Sub

Private Sub mskFI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.mskFF.SetFocus
    End If
End Sub

Private Sub mskFF_GotFocus()
    mskFF.SelStart = 0
    mskFF.SelLength = 50
End Sub

Private Sub mskFF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdLPromociones.SetFocus
    End If
End Sub

Private Function Valida(Optional pbValParametro As Boolean = True) As Boolean
    If Not IsDate(Me.mskFI.Text) Then
        MsgBox "Debe ingresar una fecha de ingreso valida.", vbInformation, "Aviso"
        Valida = False
        mskFI.SetFocus
    ElseIf Not IsDate(Me.mskFI.Text) Then
        MsgBox "Debe ingresar una fecha de ingreso valida.", vbInformation, "Aviso"
        Valida = False
        mskFI.SetFocus
    Else
        If pbValParametro Then
            If Not IsNumeric(Me.txtParametro.Text) Then
                MsgBox "Debe ingresar un paramtro valido.", vbInformation, "Aviso"
                Valida = False
                Me.txtParametro.SetFocus
            Else
                Valida = True
            End If
        Else
            Valida = True
        End If
    End If
    
End Function

Private Sub txtParametro_GotFocus()
    txtParametro.SelStart = 0
    txtParametro.SelLength = 50
End Sub

Private Sub txtParametro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdLPromociones.SetFocus
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
    End If
End Sub

Private Sub GeneraReporte()
    Dim i As Integer
    Dim k As Integer
    Dim j As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lnAcum As Currency
    
    Dim sTipoGara As String
    Dim sTipoCred As String
   
    For i = 0 To Me.Flex.Rows - 1
        lnAcum = 0
        For j = 0 To Me.Flex.Cols - 1
            xlHoja1.Cells(i + 1, j + 1) = "" & Me.Flex.TextMatrix(i, j)
        Next j
    Next i
        
    xlHoja1.Range("A1:A" & Trim(Str(Me.Flex.Rows))).Font.Bold = True
    xlHoja1.Select
    xlHoja1.Cells.EntireColumn.AutoFit
    'xlHoja1.Range("B1:B" & Trim(Str(Me.flex.Rows))).Font.Bold = True
End Sub

Private Sub cmdLPromocionesNOPF_Click()
    Dim Sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean
    Dim lnI As Integer
    Dim lsFiltro As String
    Dim lsCad As String
    Dim lsCadNoInc As String
    Dim lsConCentral As String
    Dim lsCadConsol As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    lsFiltro = ""
    For lnI = 0 To Me.lstAge.ListCount - 1
        If lstAge.Selected(lnI) Then
            If lsFiltro = "" Then
                lsFiltro = Right(lstAge.List(lnI), 2)
            Else
                lsFiltro = lsFiltro & "','" & Right(lstAge.List(lnI), 2)
            End If
        End If
    Next lnI
    
    If Not Valida(False) Then Exit Sub
    
    oCon.AbreConexion
    
    For lnI = 0 To Me.lstAge.ListCount - 1
        Sql = "  Select  PR.cPersCod cCodPers   " _
            & " From  PersPromocion PR " _
            & " Inner Join PersPromocionProd PP On PP.cPromocNro = PR.cPromocNro And cProducto = 'A2'" _
            & " Inner Join ProductoPersona PC On PR.cPersCod = PC.cPersCod And Substring(PC.cCtaCod,6,3) In ('233')" _
            & " Left  Join Captaciones As CA On CA.cCtaCod = PC.cCtaCod " _
            & " " _
            & " " _
            & " Where DateAdd(Day,-" & Me.txtParametro.Text & ",PR.dFecPromo) <= CA.dApertura And PR.dFecPromo Between '" & Format(Me.mskFI.Text, gsFormatoFecha) & "' And '" & Format(Me.mskFF.Text, gsFormatoFecha) & "'" _
            & " "
        
        If rs.State = 1 Then rs.Close
        Set rs = oCon.CargaRecordSet(Sql)
        
        While Not rs.EOF
            If lsCadNoInc = "" Then
                lsCadNoInc = rs!cCodPers
            Else
                lsCadNoInc = lsCadNoInc & "','" & rs!cCodPers
            End If
            rs.MoveNext
        Wend
        
    Next lnI
    
    
    Sql = "  Select PR.cPromocNro, PE.cPersCod, PE.cPersNombre , PR.cCodPromotor,dFecPromo, PR.cComenta, nCapMonSoles, nCapMonDolares, nColMonSoles, nColMonDolares, RTrim(ISNULL(cUbiGeoDescripcion,'')) + ' ' + PE.cPersDireccDomicilio Dir,   'DNI' + ' ' + IsNull(PEI.cPersIDnro,'') Doc, PE.cPersTelefono cTelPers  From PersPromocion PR " _
        & " Inner Join Persona PE On PE.cPersCod = PR.cPersCod " _
        & " Inner Join PersPromocionProd PP On PP.cPromocNro = PR.cPromocNro And cProducto = 'A2'" _
        & " Left Join PersID PEI On PE.cPersCod = PEI.cPersCod And cPersIDTpo = 1" _
        & " Left Join UbicacionGeografica ZO ON ZO.cUbiGeoCod = PE.cPersDireccUbiGeo" _
        & " Where PR.dFecPromo  Between '" & Format(Me.mskFI.Text, gsFormatoFecha) & "' And '" & Format(Me.mskFF.Text, gsFormatoFecha) & "' And PR.cCodAge In ('" & lsFiltro & "') And PE.cPersCod Not In ('" & lsCadNoInc & "') Order By PE.cPersNombre"
        
    'sql = " Select PR.cPromocNro, PE.cCodPers, PE.cNomPers , PR.cCodPromotor,dFecPromo, PR.cComenta, nCapMonSoles, nCapMonDolares, nColMonSoles, nColMonDolares, RTrim(ISNULL(cDesZon,'')) + ' ' + PE.cDirPers Dir,   IsNull(TC.cNomTab,'') + ' ' + IsNull(PE.cNudoci,'') Doc, PE.cTelPers  From Promocion PR" _
        & " Inner Join " & gcCentralPers & "Persona PE On PE.cCodPers = PR.cCodPers" _
        & " Inner Join PromocionProd PP On PP.cPromocNro = PR.cPromocNro And cProducto = 'A2'" _
        & " Left Join " & gcCentralCom & "Zonas ZO ON ZO.cCodZon = PE.cCodZon" _
        & " LEft Join " & gcCentralCom & "TablaCod TC ON TC.cValor = PE.cTidoci And TC.cCodTab Like '04%' " _
        & " Where PR.dFecPromo  Between '" & Format(Me.mskFI.Text, gsFormatoFecha) & "' And '" & Format(Me.mskFF.Text, gsFormatoFecha) & "' And PR.cCodAge In ('" & lsFiltro & "')" _
        & " And PE.cCodPers Not In ('" & lsCadNoInc & "') Order By PE.cNomPers"
    Set rs = oCon.CargaRecordSet(Sql)
    
    Me.Flex.Cols = 13
    Me.Flex.Rows = 1
    
    Me.Flex.TextMatrix(0, 0) = "Cod.Prom"
    Me.Flex.TextMatrix(0, 1) = "Cod.Cliente"
    Me.Flex.TextMatrix(0, 2) = "Nom.Cliente"
    Me.Flex.TextMatrix(0, 3) = "Usu.Promotor"
    Me.Flex.TextMatrix(0, 4) = "F.Promocion"
    Me.Flex.TextMatrix(0, 5) = "Comentario"
    Me.Flex.TextMatrix(0, 6) = "Monto S/. Capta"
    Me.Flex.TextMatrix(0, 7) = "Monto US$/. Capta"
    Me.Flex.TextMatrix(0, 8) = "Monto S/. Coloc"
    Me.Flex.TextMatrix(0, 9) = "Monto US$/. Coloc"
    Me.Flex.TextMatrix(0, 10) = "Direccion"
    Me.Flex.TextMatrix(0, 11) = "DNI"
    Me.Flex.TextMatrix(0, 12) = "Telefono"
    
    If Not RSVacio(rs) Then
        While Not rs.EOF
            Me.Flex.Rows = Me.Flex.Rows + 1
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 0) = "'" & rs.Fields(0)
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 1) = "'" & rs.Fields(1)
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 2) = rs.Fields(2)
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 3) = rs.Fields(3)
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 4) = Format(rs.Fields(4), ": dd/mm/yyyy")
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 5) = rs.Fields(5)
            
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 6) = Format(rs.Fields(6), "#,##0.00")
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 7) = Format(rs.Fields(7), "#,##0.00")
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 8) = Format(rs.Fields(8), "#,##0.00")
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 9) = Format(rs.Fields(9), "#,##0.00")
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 10) = rs.Fields(10)
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 11) = rs.Fields(11)
            Me.Flex.TextMatrix(Me.Flex.Rows - 1, 12) = "" & rs.Fields(12)
            
            rs.MoveNext
        Wend
    End If
    
    lsArchivoN = App.path & "\Spooler\" & Format(CDate(Me.mskFF.Text & " " & Time), gsFormatoMovFechaHora) & ".xls"
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
       If lbLibroOpen Then
          Set xlHoja1 = xlLibro.Worksheets(1)
          ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
          
          Call GeneraReporte
    
          OleExcel.Class = "ExcelWorkSheet"
          ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
          OleExcel.SourceDoc = lsArchivoN
          OleExcel.Verb = 1
          OleExcel.Action = 1
          OleExcel.DoVerb -1
       End If
       MousePointer = 0
        
    rs.Close
    Set rs = Nothing
End Sub

