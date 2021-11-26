VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmAsntoComisionBcoNac 
   Caption         =   "Comisión Banco de la Nación"
   ClientHeight    =   5625
   ClientLeft      =   465
   ClientTop       =   945
   ClientWidth     =   11970
   Icon            =   "frmAsntoComisionBcoNac.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVistaPrev 
      Caption         =   "Vista previa"
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdCargar 
      Caption         =   "Cargar"
      Height          =   315
      Left            =   4200
      TabIndex        =   10
      Top             =   240
      Width           =   1095
   End
   Begin VB.ComboBox cmbMes 
      Height          =   315
      ItemData        =   "frmAsntoComisionBcoNac.frx":030A
      Left            =   2055
      List            =   "frmAsntoComisionBcoNac.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   240
      Width           =   1980
   End
   Begin VB.TextBox txtGlosa 
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   4680
      Width           =   11055
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   11775
      Begin Sicmact.FlexEdit FEGasAge 
         Height          =   3495
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   6165
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "-Cod.Age.-Agencia-Alquiler S/.-Comisión S/.-Cta_Cont_Alquiler-Cta_Cont_Comision-ccAlqui-ccComi"
         EncabezadosAnchos=   "400-1200-3500-1500-1500-1500-1500-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-3-4-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-R-R-L-L-C-C"
         FormatosEdit    =   "0-0-0-2-2-0-0-0-0"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   10440
      TabIndex        =   0
      Top             =   5160
      Width           =   1335
   End
   Begin MSMask.MaskEdBox mskAnio 
      Height          =   300
      Left            =   585
      TabIndex        =   6
      Top             =   240
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   "_"
   End
   Begin VB.Label lblMes 
      AutoSize        =   -1  'True
      Caption         =   "Mes :"
      Height          =   195
      Left            =   1560
      TabIndex        =   9
      Top             =   285
      Width           =   390
   End
   Begin VB.Label lblAnio 
      Caption         =   "Año :"
      Height          =   210
      Left            =   120
      TabIndex        =   7
      Top             =   270
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Glosa:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   450
   End
End
Attribute VB_Name = "frmAsntoComisionBcoNac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nPost As Integer
Dim J As Integer
Dim FENoMoverdeFila As Integer
Dim rsC As ADODB.Recordset

Private Sub cmdCargar_Click()
    Dim oMov As DMov
    Set oMov = New DMov
    Dim oAje As DAgencia
    Set oAje = New DAgencia
    Dim lsFecha As String
    Dim rsMovAsnto As ADODB.Recordset
    Set rsMovAsnto = New ADODB.Recordset
    
    Dim oContFunc As NContFunciones
    Set oContFunc = New NContFunciones
    Dim nMES As Integer
    Dim nAnio As Integer
    Dim dFecha As Date

    FEGasAge.Clear
    FEGasAge.FormaCabecera
    FEGasAge.Rows = 2

    If cmbMes.ListIndex = -1 Then
       MsgBox "Seleccione un mes por favor.", vbExclamation + vbOKOnly, "Atención"
       Exit Sub
    End If

    If Val(Me.mskAnio.Text) < Year(Date) - 1 Or Val(Me.mskAnio.Text) > Year(Date) Then
       MsgBox "Ingrese un año correcto.", vbExclamation + vbOKOnly, "Atención"
       Exit Sub
    End If

    nPost = 0
    If oMov.BuscarMov(Me.mskAnio.Text & Format(cmbMes.ListIndex + 1, "00"), " cOpeCod in ('701140','701141') and nMovFlag = 0 and nMovEstado=10 ") Then
       MsgBox "Asiento ya fue generado", vbInformation, "Aviso"
       Exit Sub
    End If
    
    nMES = Val(cmbMes.ListIndex + 1) 'Month(CDate(Me.mskFecha.Text))
    nAnio = Val(Me.mskAnio.Text) 'Year(CDate(Me.mskFecha.Text))
    dFecha = DateAdd("m", 1, "01/" & Format(nMES, "00") & "/" & Format(nAnio, "0000")) - 1
    If Not oContFunc.PermiteModificarAsiento(Format(dFecha, gsFormatoMovFecha), False) Then
       Set oContFunc = Nothing
       MsgBox "Imposible generar las comisiones ya que la fecha ingresada pertenece a un mes cerrado.", vbExclamation + vbOKOnly, "Atención"
       Exit Sub
    End If
    
    Call CargarAlquiComi
End Sub

'*** PEAC 20120625
Private Sub cmdGenerar_Click()

    Dim sImp As String
    Dim sImp2 As String
    Dim J As Integer
    Dim lnMontoSuma As Double
    Dim I As Integer
    Dim nItem As Integer
    Dim oImp As New NContImprimir

    If nPost <= 0 Then
        MsgBox "No existen datos para procesar.", vbExclamation + vbOKOnly, "Atención"
        Exit Sub
    End If

    If ValidaCtasCont > 0 Then
        MsgBox "Existen Items que no tienen Ctas. Contables, verifique y créelos en el Plan Contable.", vbExclamation + vbOKOnly, "Atención"
        Exit Sub
    End If
        
    If ValidaMontos > 0 Then
        MsgBox "Existen Items que no tienen Montos, por favor verifique.", vbExclamation + vbOKOnly, "Atención"
        Exit Sub
    End If
    
    If Len(Trim(Me.txtGlosa.Text)) = 0 Then
        MsgBox "Ingrese una glosa por favor.", vbExclamation + vbOKOnly, "Atención"
        Exit Sub
    End If
    
    If MsgBox(" ¿ Está seguro de Generar el asiento contable ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmación") = vbNo Then
       Exit Sub
    End If
    Dim oMov As DMov
    Set oMov = New DMov

'    If oMov.BuscarMov(Me.mskAnio.Text & Format(cmbMes.ListIndex + 1, "00"), "cOpeCod = '" & gsOpeCod & "' and nMovEstado = '" & gMovEstContabMovContable & "' and NOT nMovFlag IN ('" & gMovFlagDeExtorno & "','" & gMovFlagExtornado & "','" & gMovFlagEliminado & "') ") Then
'       MsgBox "Asiento ya fue generado", vbInformation, "Aviso"
'       Exit Sub
'    End If

    For I = 1 To 2
    
        gsOpeCod = IIf(I = 1, "701140", "701141") ' 701140 alquiler / 701141 comision

        gdFecha = DateAdd("m", 1, CDate("01/" & Format(cmbMes.ListIndex + 1, "00") & "/" & mskAnio)) - 1
        gsMovNro = oMov.GeneraMovNro(gdFecha, gsCodAge, gsCodUser)
        
        oMov.BeginTrans
        oMov.InsertaMov gsMovNro, gsOpeCod, Trim(Me.txtGlosa.Text)
        gnMovNro = oMov.GetnMovNro(gsMovNro)
        'oMov.InsertaMovCont gnMovNro, 0, 0, ""
        nItem = 0
        lnMontoSuma = 0
        For J = 1 To nPost
            nItem = nItem + 1
            If gsOpeCod = "701140" Then
                oMov.InsertaMovCta gnMovNro, nItem, Me.FEGasAge.TextMatrix(J, 5), Me.FEGasAge.TextMatrix(J, 3)  'alquiler BN
                lnMontoSuma = lnMontoSuma + Me.FEGasAge.TextMatrix(J, 3)
            Else
                oMov.InsertaMovCta gnMovNro, nItem, Me.FEGasAge.TextMatrix(J, 6), Me.FEGasAge.TextMatrix(J, 4) 'comision BN
                lnMontoSuma = lnMontoSuma + Me.FEGasAge.TextMatrix(J, 4)
            End If
        Next
        nItem = nItem + 1
        oMov.InsertaMovCta gnMovNro, nItem, "2918090101", lnMontoSuma * -1
        
        oMov.CommitTrans
        
        If I = 1 Then
            sImp = oImp.ImprimeAsientoContable(gsMovNro, gnLinPage, gnColPage, "ASIENTO CONTABLE")
        Else
            sImp2 = oImp.ImprimeAsientoContable(gsMovNro, gnLinPage, gnColPage, "ASIENTO CONTABLE")
        End If
        
    Next
    Set oMov = Nothing
    
    sImp = sImp & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & sImp2
    EnviaPrevio sImp, Me.Caption, gnLinPage, False
    Set oImp = Nothing
    
    FEGasAge.Clear
    FEGasAge.FormaCabecera
    FEGasAge.Rows = 2
    
    Me.txtGlosa.Text = ""
    
    nPost = 0
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Function ValidaCtasCont() As Integer
    Dim lnNoCtas As Integer
    lnNoCtas = 0
    'valida si los items tienen sus ctas.cont. respectivas
    rsC.MoveFirst
    While Not rsC.EOF()
        If rsC.Fields(4) = "" Or rsC.Fields(5) = "" Then
            lnNoCtas = lnNoCtas + 1
        End If
         rsC.MoveNext
    Wend
    ValidaCtasCont = lnNoCtas
End Function

Private Function ValidaMontos() As Integer
    Dim lnNoCtas As Integer
    Dim J As Integer
    
    lnNoCtas = 0
    
    'If nPost > 0 Then
    For J = 1 To nPost
                
        If CDbl(Me.FEGasAge.TextMatrix(J, 3)) = 0 Or CDbl(Me.FEGasAge.TextMatrix(J, 4)) = 0 Then
            lnNoCtas = lnNoCtas + 1
        End If
    Next J
    
'    rsC.MoveFirst
'    While Not rsC.EOF()
'        If CDbl(rsC.Fields(2)) = 0 Or CDbl(rsC.Fields(3)) = 0 Then
'            lnNoCtas = lnNoCtas + 1
'        End If
'         rsC.MoveNext
'    Wend
    
    ValidaMontos = lnNoCtas
End Function


Private Sub cmdVistaPrev_Click()
    
    If nPost <= 0 Then
        MsgBox "No existen datos para procesar.", vbExclamation + vbOKOnly, "Atención"
        Exit Sub
    End If

    
    If ValidaCtasCont > 0 Then
        MsgBox "Existen Items que no tienen Ctas. Contables, verifique y créelos en el Plan Contable.", vbApplicationModal, "Atención"
        Exit Sub
    End If
        
    If ValidaMontos > 0 Then
        MsgBox "Existen Items que no tienen Montos, por favor verifique.", vbApplicationModal, "Atención"
        Exit Sub
    End If
        
    Dim lcTextoImp As String
    Dim lcTextoComi As String
    Dim lcTextoAlqui As String
    Dim nItem As Integer
    Dim nSumaAlqui As Double
    Dim nSumaComi As Double
    Dim I As Integer
       
    lcTextoImp = "": lcTextoComi = "": lcTextoAlqui = ""
    
    nItem = 0
    'Me.FEGasAge.TextMatrix(j, 3)
    For I = 1 To nPost
        nItem = nItem + 1
        lcTextoAlqui = lcTextoAlqui & Format(nItem, "000") & Space(1) & Left(Trim(Me.FEGasAge.TextMatrix(I, 5)) + Space(20), 20) & Left(Trim(Me.FEGasAge.TextMatrix(I, 7)) + Space(68), 68) & Right(Space(20) & Format(CDbl(Me.FEGasAge.TextMatrix(I, 3)), "#,#00.00"), 20) & oImpresora.gPrnSaltoLinea
        lcTextoComi = lcTextoComi & Format(nItem, "000") & Space(1) & Left(Trim(Me.FEGasAge.TextMatrix(I, 6)) + Space(20), 20) & Left(Trim(Me.FEGasAge.TextMatrix(I, 8)) + Space(68), 68) & Right(Space(20) & Format(CDbl(Me.FEGasAge.TextMatrix(I, 4)), "#,#00.00"), 20) & oImpresora.gPrnSaltoLinea
        nSumaAlqui = nSumaAlqui + CDbl(Me.FEGasAge.TextMatrix(I, 3))
        nSumaComi = nSumaComi + CDbl(Me.FEGasAge.TextMatrix(I, 4))
    Next I
    
'    rsC.MoveFirst
'    While Not rsC.EOF
'        nItem = nItem + 1
'        lcTextoAlqui = lcTextoAlqui & Format(nItem, "000") & Space(1) & Left(rsC.Fields(4) + Space(20), 20) & Left(rsC.Fields(6) + Space(68), 68) & Right(Space(20) & Format(rsC.Fields(2), "#,#00.00"), 20) & oImpresora.gPrnSaltoLinea
'        lcTextoComi = lcTextoComi & Format(nItem, "000") & Space(1) & Left(rsC.Fields(5) + Space(20), 20) & Left(rsC.Fields(7) + Space(68), 68) & Right(Space(20) & Format(rsC.Fields(3), "#,#00.00"), 20) & oImpresora.gPrnSaltoLinea
'        nSumaAlqui = nSumaAlqui + rsC.Fields(2)
'        nSumaComi = nSumaComi + rsC.Fields(3)
'        rsC.MoveNext
'    Wend
    nItem = nItem + 1
    lcTextoAlqui = lcTextoAlqui & Format(nItem, "000") & Space(1) & Left("2918090101" + Space(20), 20) & Left("Proveedores>>Oficina Principal" + Space(68), 68) & Space(20) & Right(Space(20) & Format(nSumaAlqui, "#,#00.00"), 20) & oImpresora.gPrnSaltoLinea
    lcTextoComi = lcTextoComi & Format(nItem, "000") & Space(1) & Left("2918090101" + Space(20), 20) & Left("Proveedores>>Oficina Principal" + Space(68), 68) & Space(20) & Right(Space(20) & Format(nSumaComi, "#,#00.00"), 20) & oImpresora.gPrnSaltoLinea
    
    
    '-------------- Alquiler
    lcTextoImp = lcTextoImp & "CAJA MUNICIPAL DE AHORRO Y CREDITO DE MAYNAS" & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & "Oficina Principal - CONTABILIDAD" & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & "A S I E N T O   C O N T A B L E" & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & "M O N E D A   N A C I O N A L" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & "Glosa: Provision Alquileres" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & "====================================================================================================================================" & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & "Itm C U E N T A   C O N T A B L E                                                                           DEBE            HABER" & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & "    Codigo Descripcion" & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & "------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & "EN MONEDA NACIONAL" & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & "--------------------" & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & lcTextoAlqui
    lcTextoImp = lcTextoImp & "------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & Space(92) & Right(Space(20) & Format(nSumaAlqui, "#,#00.00"), 20) & Right(Space(20) & Format(nSumaAlqui, "#,#00.00"), 20) & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & "====================================================================================================================================" & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    
    '-------------- Comision
    lcTextoImp = lcTextoImp & "CAJA MUNICIPAL DE AHORRO Y CREDITO DE MAYNAS" & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & "Oficina Principal - CONTABILIDAD" & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & "A S I E N T O   C O N T A B L E" & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & "M O N E D A   N A C I O N A L" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & "Glosa: Provision Comision" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & "====================================================================================================================================" & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & "Itm C U E N T A   C O N T A B L E                                                                           DEBE            HABER" & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & "    Codigo Descripcion" & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & "------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & "EN MONEDA NACIONAL" & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & "--------------------" & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & lcTextoComi
    lcTextoImp = lcTextoImp & "------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & Space(92) & Right(Space(20) & Format(nSumaComi, "#,#00.00"), 20) & Right(Space(20) & Format(nSumaComi, "#,#00.00"), 20) & oImpresora.gPrnSaltoLinea
    lcTextoImp = lcTextoImp & "====================================================================================================================================" & oImpresora.gPrnSaltoLinea
    
    '----------------
    
    EnviaPrevio lcTextoImp, "Asiento Contable", 1, False

End Sub

Private Sub Command3_Click()
    
  
    
'    Dim obDAgencia As DAgencia
'    Set obDAgencia = New DAgencia
'    If nPost > 0 Then
'    For j = 1 To nPost
'
'        obDAgencia.ActualizarAgenciaPorcentajeSeguPatri FEGasAge.TextMatrix(j, 1), FEGasAge.TextMatrix(j, 3), FEGasAge.TextMatrix(j, 5)
'
'    Next j
'    MsgBox "Datos se registraron correctamente", vbApplicationModal
'    End If

End Sub

Private Sub FEGasAge_click()
   Call FEGasAge_KeyPress(13)
End Sub

Private Sub FEGasAge_KeyPress(KeyAscii As Integer)
 FENoMoverdeFila = FEGasAge.Row
 FEGasAge.lbEditarFlex = True
End Sub

Private Sub Form_Load()
    
'    Dim obDAgencia As DAgencia
'    Set obDAgencia = New DAgencia
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    Set rs = obDAgencia.GetAgenciaPorcentajeSeguPatri
'    Call cmdCargarArch(rs)
    
    Dim oGen As DGeneral
    Set oGen = New DGeneral
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    
    Set rs1 = oGen.GetConstante(1010)
    Me.cmbMes.Clear
    While Not rs1.EOF
        cmbMes.AddItem rs1.Fields(0) & Space(50) & rs1.Fields(1)
        rs1.MoveNext
    Wend
    
    Me.mskAnio.Text = Format(gdFecSis, "yyyy")
    
End Sub

Private Sub CargarAlquiComi()
    
    Dim lsFecha, lsTipo As String
    Dim obDAgencia As DAgencia
    Set obDAgencia = New DAgencia
           
    Set rsC = New ADODB.Recordset
    
    lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    
    If Len(Trim(Right(cmbMes.Text, 2))) = 1 Then
        lsFecha = mskAnio.Text & "0" & Trim(Right(cmbMes.Text, 2)) + "01"
    Else
        lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2)) + "01"
    End If
    
    Set rsC = obDAgencia.ObtieneComiBcoNac(lsFecha)
    
    Dim I As Integer
    If nPost > 0 Then
        For I = 1 To nPost
            FEGasAge.EliminaFila (1)
        Next I
    End If
    nPost = 0
    If (rsC.EOF And rsC.BOF) Then
        MsgBox "No existen comisiones del Banco de la Nación"
        Exit Sub
    End If
    rsC.MoveFirst

    nPost = 0
    Do While Not rsC.EOF
        nPost = nPost + 1
        FEGasAge.AdicionaFila
        FEGasAge.TextMatrix(nPost, 0) = nPost
        FEGasAge.TextMatrix(nPost, 1) = IIf(rsC!cAgecod = "", "-", rsC!cAgecod)
        FEGasAge.TextMatrix(nPost, 2) = rsC!cAgeDescripcion
        FEGasAge.TextMatrix(nPost, 3) = Format(rsC!Alquiler, "#0.00")
        FEGasAge.TextMatrix(nPost, 4) = Format(rsC!Comision, "#0.00")
        FEGasAge.TextMatrix(nPost, 5) = rsC!CtaContAlqui
        FEGasAge.TextMatrix(nPost, 6) = rsC!CtaContComi
        FEGasAge.TextMatrix(nPost, 7) = rsC!ccAlqui
        FEGasAge.TextMatrix(nPost, 8) = rsC!ccComi
        rsC.MoveNext
    Loop
        
End Sub


'Private Sub cmdCargarArch(rs As ADODB.Recordset)
'
'        Dim i As Integer
'        If nPost > 0 Then
'            For i = 1 To nPost
'                FEGasAge.EliminaFila (1)
'            Next i
'        End If
'        nPost = 0
'        If (rs.EOF And rs.BOF) Then
'            MsgBox "No existen porcenctajes de gastos de Agencias"
'            Exit Sub
'        End If
'        rs.MoveFirst
'
'        nPost = 0
'        'Do While Not (rs.EOF Or rs.BOF)
'        Do While Not rs.EOF '' Or rs.BOF)
'            nPost = nPost + 1
'            FEGasAge.AdicionaFila
'            FEGasAge.TextMatrix(nPost, 0) = nPost
'            FEGasAge.TextMatrix(nPost, 1) = IIf(rs!cAgecod = "", "-", rs!cAgecod)
'            FEGasAge.TextMatrix(nPost, 2) = rs!cAgeDescripcion
'            FEGasAge.TextMatrix(nPost, 3) = rs!nTipoSeguro
'            FEGasAge.TextMatrix(nPost, 4) = IIf(rs!cConsDescripcion = "", "-", rs!cConsDescripcion)
'            FEGasAge.TextMatrix(nPost, 5) = Format(rs!nAgePorcentaje, "#0.00")
'            rs.MoveNext
'        Loop
'
'End Sub
