VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmResumenCredPignoRecup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen de Créditos Pignoraticios Recuperados"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12585
   Icon            =   "frmResumenCredPignoRecup.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   12585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExtorna 
      Caption         =   "Extorna Asiento"
      Height          =   345
      Left            =   2520
      TabIndex        =   13
      Top             =   6360
      Width           =   1560
   End
   Begin VB.ComboBox cboTpo 
      Height          =   315
      Left            =   5055
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   0
      Width           =   2520
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Height          =   345
      Left            =   45
      TabIndex        =   8
      Top             =   6375
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar Asiento"
      Height          =   345
      Left            =   4215
      TabIndex        =   7
      Top             =   6360
      Width           =   1440
   End
   Begin VB.ComboBox cmbMes 
      Height          =   315
      ItemData        =   "frmResumenCredPignoRecup.frx":030A
      Left            =   1935
      List            =   "frmResumenCredPignoRecup.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   30
      Width           =   2220
   End
   Begin MSMask.MaskEdBox mskAnio 
      Height          =   300
      Left            =   600
      TabIndex        =   4
      Top             =   60
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   9375
      TabIndex        =   2
      Top             =   6375
      Width           =   960
   End
   Begin VB.CommandButton cmdDeprecia 
      Caption         =   "&Generar Cálculo"
      Height          =   345
      Left            =   7740
      TabIndex        =   1
      Top             =   6360
      Width           =   1560
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   345
      Left            =   6690
      TabIndex        =   0
      Top             =   6375
      Width           =   960
   End
   Begin Sicmact.FlexEdit FlexEdit1 
      Height          =   5775
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   10186
      Cols0           =   19
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   $"frmResumenCredPignoRecup.frx":030E
      EncabezadosAnchos=   "300-700-2000-1200-1200-1200-1200-1200-1200-1200-1200-1200-1000-1000-1000-1000-1200-1200-1200"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-L-R-R-R-R-R-R-R-L-R-R-R-R-R-R-R"
      FormatosEdit    =   "0-0-0-1-2-2-2-2-2-2-2-1-2-2-2-3-2-2-2"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      Appearance      =   0
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Label lblTipo 
      AutoSize        =   -1  'True
      Caption         =   "Agencias :"
      Height          =   195
      Left            =   4215
      TabIndex        =   11
      Top             =   60
      Width           =   750
   End
   Begin VB.OLE OleExcel 
      Appearance      =   0  'Flat
      AutoActivate    =   3  'Automatic
      Enabled         =   0   'False
      Height          =   240
      Left            =   1215
      SizeMode        =   1  'Stretch
      TabIndex        =   9
      Top             =   6420
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblMes 
      Caption         =   "Mes :"
      Height          =   210
      Left            =   1440
      TabIndex        =   6
      Top             =   75
      Width           =   510
   End
   Begin VB.Label lblAnio 
      Caption         =   "Año :"
      Height          =   210
      Left            =   135
      TabIndex        =   3
      Top             =   90
      Width           =   705
   End
End
Attribute VB_Name = "frmResumenCredPignoRecup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsCaption As String
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim nMes As Integer
Dim nAnio As Integer
Dim dFecha As Date

Public Sub Ini(psCaption As String)
    lsCaption = psCaption
    Me.Show 1
End Sub



Private Sub cmdDeprecia_Click()
    
    If Not IsNumeric(Me.mskAnio.Text) Then
        MsgBox "Debe Ingresar un año Valido.", vbInformation, "Aviso"
        Me.mskAnio.SetFocus
        Exit Sub
    ElseIf Me.cmbMes.Text = "" Then
        MsgBox "Debe Ingresar un mes Valido.", vbInformation, "Aviso"
        Me.cmbMes.SetFocus
        Exit Sub
    ElseIf Me.cboTpo.Text = "" Then
        MsgBox "Indique alguna agencia o todas.", vbInformation, "Aviso"
        Me.cboTpo.SetFocus
        Exit Sub
    End If
       
    Call llenagrid

End Sub

Private Sub cmdExtorna_Click()
    Dim oContFunc As NContFunciones
    Dim oAge As DOperacion
    Dim overi As DOperacion
    Dim rs As ADODB.Recordset
    Dim lcFecPeriodo As String
    Dim lcMovNroExtorno As Integer
    Dim lnMovNroExtorno As Integer
    Dim lnMovNroExt As Long
    Dim lcTextoMov As String
    Dim lcTextoAsientoExtor As String
    Dim nFlag As Integer
    Dim lnI As Integer
    Dim lcAgeActual As String
    
    lcFecPeriodo = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    If Len(Trim(Right(cmbMes.Text, 2))) = 1 Then
        lcFecPeriodo = mskAnio.Text & "0" & Trim(Right(cmbMes.Text, 2))
    Else
        lcFecPeriodo = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    End If
    
    gsOpeCod = "300460"
    
    nMes = Val(Trim(Right(cmbMes.Text, 2)))
    nAnio = Val(mskAnio.Text)
    dFecha = DateAdd("m", 1, "01/" & Format(nMes, "00") & "/" & Format(nAnio, "0000")) - 1
    Set oContFunc = New NContFunciones
    If Not oContFunc.PermiteModificarAsiento(Format(dFecha, gsFormatoMovFecha), False) Then
       Set oContFunc = Nothing
       MsgBox "Imposible realizar este proceso ya que la fecha ingresada pertenece a un mes cerrado.", vbInformation, "Aviso"
       Exit Sub
    End If
    
    nFlag = 0
    
    If MsgBox("¿Desea extornar el asiento contable? ", vbQuestion + vbYesNo, "Atención") = vbNo Then Exit Sub
        
     lcAgeActual = ""
    For lnI = 1 To Me.FlexEdit1.Rows - 1
        If lnI = 1 Then
            lcAgeActual = Me.FlexEdit1.TextMatrix(lnI, 1)
        End If

        If Me.FlexEdit1.TextMatrix(lnI, 1) = lcAgeActual Then
            lcAgeActual = Me.FlexEdit1.TextMatrix(lnI, 1)
        Else

            If Me.FlexEdit1.TextMatrix(lnI, 1) = "" Then
        
                Set overi = New DOperacion
                Set rs = overi.VerificaAsientoCont(gsOpeCod, lcFecPeriodo, lcAgeActual)
                Set overi = Nothing
            
                If (rs.EOF And rs.BOF) Then
                    lcTextoMov = lcTextoMov + lcAgeActual + "-"
                Else
                    lnMovNroExt = rs!nMovNro
                    lcTextoAsientoExtor = lcTextoAsientoExtor + lcAgeActual + "-"
                End If
                nFlag = 1
                lcAgeActual = ""
                Set oAge = New DOperacion
                Call oAge.EliminaAsientoCont(lnMovNroExt)
                Set oAge = Nothing
                lcAgeActual = Me.FlexEdit1.TextMatrix(lnI, 1)
            Else
                lcAgeActual = Me.FlexEdit1.TextMatrix(lnI, 1)
            End If
        End If
    Next
    
    If nFlag = 1 Then
        MsgBox "Se extornaron los asientos contables de las siguientes agencias: " + lcTextoAsientoExtor, vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdGrabar_Click()
    
    Dim oMov As DMov
    Set oMov = New DMov
    
    Dim oDep As DOperacion
    Set oDep = New DOperacion
       
    Dim oConect As DConecta
    Set oConect = New DConecta
    
    Dim lnMovNro As Long
    Dim lsMovNro As String
    
    Dim lsTipo As String
    Dim lsFecha As String
    Dim I As Integer
    Dim lnI As Long
    Dim lnContador As Long
    Dim lsCtaCont As String
    Dim oPrevio As clsPrevioFinan
    Dim oAsiento As NContImprimir
    Dim nConta As Integer, lcCtaDif As String
    Dim overi As DOperacion
    Dim lnDebe As Double, lnHaber As Double, lnTotHaber As Double, lnTotDebe As Double
    Dim lnDebeME As Double, lnTotDebeME As Double
    Dim oContFunc As NContFunciones
    Dim lnMontoPrin As Currency
    Dim rsAgesDistrib As ADODB.Recordset
    Dim lsSql As String, lcCtaCont As String
    Dim rsBuscaCuenta As ADODB.Recordset
    Dim lnItemDistri As Integer
    Dim lnRegImporte As Currency
    Dim lnItemPrin As Integer
    Dim lnImpoPrin As Currency
    Dim lnMontoPrME As Currency
    Dim lnRegImporMETot As Currency
    Dim lnMonto As Currency
    Dim lsCtasInexis As String
    Dim lnRegImporME As Currency
    Dim lnRegImporteTot As Currency
    Dim lnImpoPrinME As Currency
    Dim lnMontoDebePrin As Currency
    Set oPrevio = New clsPrevioFinan
    Set oAsiento = New NContImprimir
   
    Dim rs As ADODB.Recordset, rs1 As ADODB.Recordset
    
    Dim lcAgeActual As String
    
    Dim ldFechaDepre As Date
    Dim ldFechaRegistro As Date
    
    Dim lnAdjudi  As Double, lnTotProv As Double, lnValorNeto As Double
    Dim lnAdjudiBaja As Double, lnTotProvBaja As Double 'PASI20140811 TI-ERS107-2014 ,Agregado para generar el asiento contable de Baja de Bienes Adjudicados por la venta
    Dim lcCtaHAdju As String, lcCtaDTotProv As String, lcCtaDValorNeto As String
   Dim lcUltAgeReg As String
   Dim lnFlag As Integer
   Dim lnItem As Integer
   Dim lcMovNro As String
   Dim lsGlosa As String
   Dim txtMeses As String
   
If Me.FlexEdit1.Rows - 2 = 0 Then
    MsgBox "Genere el cálculo por favor."
    Exit Sub
End If

gsOpeCod = "300460"
lcAgeActual = ""

lnAdjudi = 0
lnTotProv = 0
lnValorNeto = 0

'PASI20140811 TI-ERS107-2014
lnAdjudiBaja = 0
lnTotProvBaja = 0

'end PASI

lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))

If Len(Trim(Right(cmbMes.Text, 2))) = 1 Then
    lsFecha = mskAnio.Text & "0" & Trim(Right(cmbMes.Text, 2))
Else
    lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
End If

If MsgBox("¿Desea grabar el asiento? ", vbQuestion + vbYesNo, "Atención") = vbNo Then Exit Sub

lnFlag = 0
lcUltAgeReg = ""
For lnI = 1 To Me.FlexEdit1.Rows - 1

    If lnI = 1 Then
        lcAgeActual = Me.FlexEdit1.TextMatrix(lnI, 1)
    End If
    
    If Me.FlexEdit1.TextMatrix(lnI, 1) = lcAgeActual Then
        'Modificado PASI20140811 TI-ERS107-2014
        'lnAdjudi = lnAdjudi + Me.FlexEdit1.TextMatrix(lnI, 13)
        'lnTotProv = lnTotProv + Me.FlexEdit1.TextMatrix(lnI, 17)
        'lnValorNeto = lnValorNeto + Me.FlexEdit1.TextMatrix(lnI, 18)
        
        'lcAgeActual = Me.FlexEdit1.TextMatrix(lnI, 1)
        
        If CInt(Me.FlexEdit1.TextMatrix(lnI, 15)) >= 18 Then
            lnAdjudiBaja = lnAdjudiBaja + Me.FlexEdit1.TextMatrix(lnI, 13)
            lnTotProvBaja = lnTotProvBaja + Me.FlexEdit1.TextMatrix(lnI, 17)
        Else
            lnAdjudi = lnAdjudi + Me.FlexEdit1.TextMatrix(lnI, 13)
            lnTotProv = lnTotProv + Me.FlexEdit1.TextMatrix(lnI, 17)
            lnValorNeto = lnValorNeto + Me.FlexEdit1.TextMatrix(lnI, 18)
        End If
        lcAgeActual = Me.FlexEdit1.TextMatrix(lnI, 1)
        'end PASI
    Else
    
        Set rs = New ADODB.Recordset
        
        Set overi = New DOperacion
        Set rs = overi.VerificaAsientoCont(gsOpeCod, lsFecha, lcAgeActual)
        Set rs1 = overi.ObtieneCtasResumenCredPigno
        Set overi = Nothing
        
        If Not rs.EOF Then
            If lcUltAgeReg <> lcAgeActual And lsCtasInexis <> lcAgeActual Then
                lsCtasInexis = lsCtasInexis + lcAgeActual + "-"
            End If
        Else
            If Abs(lnAdjudi + lnTotProv + lnValorNeto) > 0 Or Abs(lnAdjudiBaja + lnTotProvBaja) > 0 Then
                
                lnItem = 0
            
                nMes = Val(Trim(Right(cmbMes.Text, 2)))
                nAnio = Val(mskAnio.Text)
                dFecha = DateAdd("m", 1, "01/" & Format(nMes, "00") & "/" & Format(nAnio, "0000")) - 1
                Set oContFunc = New NContFunciones
                If Not oContFunc.PermiteModificarAsiento(Format(dFecha, gsFormatoMovFecha), False) Then
                   Set oContFunc = Nothing
                   MsgBox "Imposible grabar el asiento en un mes cerrado.", vbInformation, "Aviso"
                   Exit Sub
                End If
                ldFechaRegistro = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & Right(Me.cmbMes.Text, 2) & "/" & Me.mskAnio.Text)))
                
                oMov.BeginTrans
            
                    lcUltAgeReg = lcAgeActual
                    
                    
                    If Abs(lnAdjudi + lnTotProv + lnValorNeto) > 0 Then 'Agregado PASI20140811 TI-ERS107-2014
                        
                        'lsMovNro = oMov.GeneraMovNro(ldFechaRegistro, Right(gsCodAge, 2), gsCodUser)
                        lsMovNro = oMov.GeneraMovNro(ldFechaRegistro, lcAgeActual, gsCodUser)
                        
                        '***PEAC 20111102
                        
                        txtMeses = Choose(Mid(lsMovNro, 5, 2), "Enero", "Febrero", "Marzo", "Abril", _
                                                "Mayo", "Junio", "Julio", "Agosto", _
                                                "Setiembre", "Octubre", "Noviembre", "Diciembre")
           
                        lsGlosa = "Joyas Adjudicadas y vendidas en el mes de " & txtMeses & " del " & Left(lsMovNro, 4)
                        
                        'oMov.InsertaMov lsMovNro, gsOpeCod, "REG. " & Trim(Mid(Me.cboTpo.Text, 1, Len(Me.cboTpo.Text) - 2))
                        oMov.InsertaMov lsMovNro, gsOpeCod, lsGlosa
                        
                        '*** FIN PEAC
                        
                        lnMovNro = oMov.GetnMovNro(lsMovNro)
                    
                        lcMovNro = lcMovNro + "'" + lsMovNro + "',"
                    
                        lnFlag = 1
                    
                        lcCtaHAdju = Replace(rs1!HAdju, "AG", lcAgeActual)
                        lcCtaDTotProv = Replace(rs1!DTotProv, "AG", lcAgeActual)
                        lcCtaDValorNeto = Replace(rs1!DValorNeto, "AG", lcAgeActual)
        
                        lnItem = lnItem + 1
                        oMov.InsertaMovCta lnMovNro, lnItem, lcCtaHAdju, lnAdjudi * -1
                        lnItem = lnItem + 1
                        oMov.InsertaMovCta lnMovNro, lnItem, lcCtaDTotProv, lnTotProv
                        lnItem = lnItem + 1
                        oMov.InsertaMovCta lnMovNro, lnItem, lcCtaDValorNeto, lnValorNeto
                    End If
                    'Agregado PASI20140811 TI-ERS107-2014
                    If Abs(lnAdjudiBaja + lnTotProvBaja) > 0 Then
                        lnItem = 0
                        Sleep 1000
                        lsMovNro = oMov.GeneraMovNro(ldFechaRegistro, lcAgeActual, gsCodUser)
                        txtMeses = Choose(Mid(lsMovNro, 5, 2), "Enero", "Febrero", "Marzo", "Abril", _
                                                "Mayo", "Junio", "Julio", "Agosto", _
                                                "Setiembre", "Octubre", "Noviembre", "Diciembre")
                        lsGlosa = "Joyas Adjudicadas y vendidas en el mes de " & txtMeses & " del " & Left(lsMovNro, 4)
                        oMov.InsertaMov lsMovNro, gsOpeCod, lsGlosa
                        lnMovNro = oMov.GetnMovNro(lsMovNro)
                        lcMovNro = lcMovNro + "'" + lsMovNro + "',"
                        lnFlag = 1
                        lcCtaHAdju = Replace(rs1!HAdju, "AG", lcAgeActual)
                        lcCtaDTotProv = Replace(rs1!DTotProv, "AG", lcAgeActual)
                        lnItem = lnItem + 1
                        oMov.InsertaMovCta lnMovNro, lnItem, lcCtaHAdju, lnAdjudiBaja * -1
                        lnItem = lnItem + 1
                        oMov.InsertaMovCta lnMovNro, lnItem, lcCtaDTotProv, lnTotProvBaja
                    End If
                    'end PASI
                oMov.CommitTrans
            End If
        End If
    
        lnAdjudi = 0
        lnTotProv = 0
        lnValorNeto = 0
        
        'Agregado PASI20140811 TI-ERS107-2014
        lnAdjudiBaja = 0
        lnTotProvBaja = 0
        'end PASI
        
        If Me.FlexEdit1.TextMatrix(lnI, 1) <> "" Then
            'Modificado PASI20140811 TI-ERS107-2014
            'lnAdjudi = lnAdjudi + Me.FlexEdit1.TextMatrix(lnI, 13)
            'lnTotProv = lnTotProv + Me.FlexEdit1.TextMatrix(lnI, 17)
            'lnValorNeto = lnValorNeto + Me.FlexEdit1.TextMatrix(lnI, 18)
            
            'lcAgeActual = Me.FlexEdit1.TextMatrix(lnI, 1)
            
            If CInt(Me.FlexEdit1.TextMatrix(lnI, 15)) >= 18 Then
                lnAdjudiBaja = lnAdjudiBaja + Me.FlexEdit1.TextMatrix(lnI, 13)
                lnTotProvBaja = lnTotProvBaja + Me.FlexEdit1.TextMatrix(lnI, 17)
            Else
                lnAdjudi = lnAdjudi + Me.FlexEdit1.TextMatrix(lnI, 13)
                lnTotProv = lnTotProv + Me.FlexEdit1.TextMatrix(lnI, 17)
                lnValorNeto = lnValorNeto + Me.FlexEdit1.TextMatrix(lnI, 18)
            End If
            lcAgeActual = Me.FlexEdit1.TextMatrix(lnI, 1)
            'end PASI
        End If
    End If
Next

    If lsCtasInexis <> "" Then
        MsgBox "Las siguientes agencias ya tienen asientos generados en este mes y año.: " + Chr(10) + lsCtasInexis, vbOKOnly, "Aviso"
    End If
    
    If lnFlag = 1 Then
    
        lcMovNro = Left(lcMovNro, Len(lcMovNro) - 1) + IIf(Right(lcMovNro, 1) = ",", "", "")
        oPrevio.Show oAsiento.ImprimeAsientoContResVtaPigno(lcMovNro, 60, 80, "A S I E N T O  C O N T A B L E - " + Trim(Left(Me.cboTpo.Text, 30)) + " - " + CStr(dFecha)), "", True
    End If

End Sub

Private Sub cmdImprimir_Click()
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean
    
    If Me.FlexEdit1.TextMatrix(1, 1) = "" Then
        MsgBox "Debe Depreciar antes de imprimir.", vbInformation, "Aviso"
        Me.cmdDeprecia.SetFocus
        Exit Sub
    End If
    
    lsArchivoN = App.path & "\Spooler\" & Format(gdFecSis, "yyyymmdd") & ".xls"
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
       Call GeneraReporte(Me.FlexEdit1.GetRsNew)
       OleExcel.Class = "ExcelWorkSheet"
       ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
       OleExcel.SourceDoc = lsArchivoN
       OleExcel.Verb = 1
       OleExcel.Action = 1
       OleExcel.DoVerb -1
    End If
    MousePointer = 0
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()

    Dim overi As DOperacion
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim oGen As DGeneral
    Set oGen = New DGeneral
    
    Set rs = oGen.GetConstante(1010)
    Me.cmbMes.Clear
    While Not rs.EOF
        cmbMes.AddItem rs.Fields(0) & Space(50) & rs.Fields(1)
        rs.MoveNext
    Wend
    
    
    Set overi = New DOperacion
    Set rs = overi.ObtieneAgencias()
    Set overi = Nothing

    Me.cboTpo.Clear
    While Not rs.EOF
        cboTpo.AddItem rs.Fields(1) & Space(50) & rs.Fields(0)
        rs.MoveNext
    Wend

      
    Me.mskAnio.Text = Format(gdFecSis, "yyyy")
     

End Sub

Private Sub GeneraReporte(prRs As ADODB.Recordset)
    Dim I As Integer
    Dim K As Integer
    Dim j As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lsSuma As String
    Dim sTipoGara As String
    Dim sTipoCred As String
    Dim lnAcum As Variant 'Currency
    Dim lnSer As Currency
    
    I = -1
    prRs.MoveFirst
    I = I + 1
    For j = 0 To prRs.Fields.Count - 1
        xlHoja1.Cells(I + 1, j + 1) = prRs.Fields(j).Name
    Next j
    
    While Not prRs.EOF
        If Len(prRs.Fields(0)) > 0 Then
            I = I + 1
            For j = 0 To prRs.Fields.Count - 1
                xlHoja1.Cells(I + 1, j + 1) = prRs.Fields(j)
            Next j
        End If
        prRs.MoveNext
    Wend

    I = I + 1
'    xlHoja1.Range("A1:A" & Trim(Str(I))).Font.Bold = True
'    xlHoja1.Range("B1:B" & Trim(Str(I))).Font.Bold = True
    xlHoja1.Range("1:1").Font.Bold = True

    xlHoja1.Range("C1:C" & Trim(Str(I))).NumberFormat = "dd/mm/yyyy"

    xlHoja1.Range("D1:D" & Trim(Str(I))).NumberFormat = "#,##0.00"
    xlHoja1.Range("E1:E" & Trim(Str(I))).NumberFormat = "#,##0.00"
    xlHoja1.Range("F1:F" & Trim(Str(I))).NumberFormat = "#,##0.00"
    xlHoja1.Range("G1:G" & Trim(Str(I))).NumberFormat = "#,##0.00"
    xlHoja1.Range("H1:H" & Trim(Str(I))).NumberFormat = "#,##0.00"
    xlHoja1.Range("I1:I" & Trim(Str(I))).NumberFormat = "#,##0.00"
    xlHoja1.Range("J1:J" & Trim(Str(I))).NumberFormat = "#,##0.00"
    
    xlHoja1.Range("K1:K" & Trim(Str(I))).NumberFormat = "dd/mm/yyyy"
    
    xlHoja1.Range("L1:L" & Trim(Str(I))).NumberFormat = "#,##0.00"
    xlHoja1.Range("M1:M" & Trim(Str(I))).NumberFormat = "#,##0.00"
    xlHoja1.Range("N1:N" & Trim(Str(I))).NumberFormat = "#,##0.00"
    xlHoja1.Range("O1:O" & Trim(Str(I))).NumberFormat = "#,##0.00"
    xlHoja1.Range("P1:P" & Trim(Str(I))).NumberFormat = "#,##0.00"
    xlHoja1.Range("Q1:Q" & Trim(Str(I))).NumberFormat = "#,##0.00"
    xlHoja1.Range("R1:R" & Trim(Str(I))).NumberFormat = "#,##0.00"
       
    
    xlHoja1.Range("D" & Trim(Str(I + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(I)) & "]C:R[-1]C)"
    xlHoja1.Range("E" & Trim(Str(I + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(I)) & "]C:R[-1]C)"
    xlHoja1.Range("F" & Trim(Str(I + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(I)) & "]C:R[-1]C)"
    xlHoja1.Range("G" & Trim(Str(I + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(I)) & "]C:R[-1]C)"
    xlHoja1.Range("H" & Trim(Str(I + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(I)) & "]C:R[-1]C)"
    xlHoja1.Range("I" & Trim(Str(I + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(I)) & "]C:R[-1]C)"
    xlHoja1.Range("J" & Trim(Str(I + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(I)) & "]C:R[-1]C)"
    
    xlHoja1.Range("L" & Trim(Str(I + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(I)) & "]C:R[-1]C)"
    xlHoja1.Range("M" & Trim(Str(I + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(I)) & "]C:R[-1]C)"
    xlHoja1.Range("N" & Trim(Str(I + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(I)) & "]C:R[-1]C)"
    
    xlHoja1.Range("P" & Trim(Str(I + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(I)) & "]C:R[-1]C)"
    xlHoja1.Range("Q" & Trim(Str(I + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(I)) & "]C:R[-1]C)"
    xlHoja1.Range("R" & Trim(Str(I + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(I)) & "]C:R[-1]C)"
    
    xlHoja1.Columns.AutoFit


End Sub

Private Sub GeneraReporteSaldoHistorico(prRs As ADODB.Recordset)
    Dim I As Integer
    Dim K As Integer
    Dim j As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lsSuma As String
    Dim sTipoGara As String
    Dim sTipoCred As String
    Dim lnAcum As Currency
    
    I = 1
    
    With xlHoja1.Range("A1:N1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlHoja1.Range("A1:N1").Merge
    xlHoja1.Range("A1:N1").FormulaR1C1 = " REPORTE DE SALDOS HISTORICOS "
    
    With xlHoja1.Range("A2:G2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlHoja1.Range("A2:G2").Merge
    xlHoja1.Range("A2:G2").FormulaR1C1 = " COSTO DEL ACTIVO FIJO "
    
    With xlHoja1.Range("H2:N2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlHoja1.Range("H2:N2").Merge
    xlHoja1.Range("H2:N2").FormulaR1C1 = " DEPRECIACION DE ACTIVOS FIJOS "
    
    prRs.MoveFirst
    While Not prRs.EOF
        I = I + 1
        
        If I = 2 Then
            xlHoja1.Cells(I + 1, 1) = "COD CTIVO"
            xlHoja1.Cells(I + 1, 2) = "COD. PATRIM."
            xlHoja1.Cells(I + 1, 3) = "F. ADQUIS"
            xlHoja1.Cells(I + 1, 4) = "SALDO AÑO ANT."
            xlHoja1.Cells(I + 1, 5) = "COMPRAS AÑO"
            xlHoja1.Cells(I + 1, 6) = "RETIROS AÑO"
            xlHoja1.Cells(I + 1, 7) = "SALDO ACTUAL"
            xlHoja1.Cells(I + 1, 8) = "DEP ACUM EJER ANT"
            xlHoja1.Cells(I + 1, 9) = "DEP AL MES ANT"
            xlHoja1.Cells(I + 1, 10) = "DEP DEP MES"
            xlHoja1.Cells(I + 1, 11) = "TOT DEP DEL EJER"
            xlHoja1.Cells(I + 1, 12) = "DEP ACUM DE RET"
            xlHoja1.Cells(I + 1, 13) = "DEP ACUM TOTAL"
            xlHoja1.Cells(I + 1, 14) = "VALOR EN LIBROS"
            
            I = I + 1
            
            xlHoja1.Cells(I + 1, 1) = prRs!Codigo & "-" & prRs!Serie
            xlHoja1.Cells(I + 1, 2) = Format(I - 2, "00000")
            xlHoja1.Cells(I + 1, 3) = "'" & Format(prRs!F_Compra, gsFormatoFechaView)
            xlHoja1.Cells(I + 1, 4) = Format(IIf(Year(prRs!F_Compra) <> Me.mskAnio.Text, prRs!Valor, 0), "#,##0.00")
            xlHoja1.Cells(I + 1, 5) = Format(IIf(Year(prRs!F_Compra) = Me.mskAnio.Text, prRs!Valor, 0), "#,##0.00")
            
            If prRs!Baja = "False" Then
                xlHoja1.Cells(I + 1, 6) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(I + 1, 6) = Format(prRs!Valor, "#,##0.00")
                Else
                    xlHoja1.Cells(I + 1, 6) = Format(0, "#,##0.00")
                End If
            End If
            
            xlHoja1.Cells(I + 1, 7) = Format(xlHoja1.Cells(I + 1, 4) + xlHoja1.Cells(I + 1, 5) - xlHoja1.Cells(I + 1, 6), "#,##0.00")
            
            xlHoja1.Cells(I + 1, 8) = Format(IIf(Year(prRs!F_Compra) <> Me.mskAnio.Text, prRs!Dep_H_Ejer_Ant, 0), "#,##0.00")
            xlHoja1.Cells(I + 1, 9) = IIf(prRs!Baja = "False", Format(prRs!Dep_H_Mes_Ant, "#,##0.00"), Format(0, "#,##0.00"))
            xlHoja1.Cells(I + 1, 10) = Format(prRs!Dep_H_Mes, "#,##0.00")
            xlHoja1.Cells(I + 1, 11) = Format(xlHoja1.Cells(I + 1, 9) + xlHoja1.Cells(I + 1, 10), "#,##0.00")
            If prRs!Baja = "False" Then
                xlHoja1.Cells(I + 1, 12) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(I + 1, 12) = Format(prRs!Dep_H_Mes_Ant, "#,##0.00")
                Else
                    xlHoja1.Cells(I + 1, 12) = Format(0, "#,##0.00")
                End If
            End If
            xlHoja1.Cells(I + 1, 13) = Format(xlHoja1.Cells(I + 1, 8) + xlHoja1.Cells(I + 1, 11) - xlHoja1.Cells(I + 1, 12), "#,##0.00")
            xlHoja1.Cells(I + 1, 14) = Format(xlHoja1.Cells(I + 1, 13) - xlHoja1.Cells(I + 1, 7), "#,##0.00")
        Else
            xlHoja1.Cells(I + 1, 1) = prRs!Codigo & "-" & prRs!Serie
            xlHoja1.Cells(I + 1, 2) = Format(I - 2, "00000")
            xlHoja1.Cells(I + 1, 3) = "'" & Format(prRs!F_Compra, gsFormatoFechaView)
            xlHoja1.Cells(I + 1, 4) = Format(IIf(Year(prRs!F_Compra) <> Me.mskAnio.Text, prRs!Valor, 0), "#,##0.00")
            xlHoja1.Cells(I + 1, 5) = Format(IIf(Year(prRs!F_Compra) = Me.mskAnio.Text, prRs!Valor, 0), "#,##0.00")
            
            If prRs!Baja = "False" Then
                xlHoja1.Cells(I + 1, 6) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(I + 1, 6) = Format(prRs!Valor, "#,##0.00")
                Else
                    xlHoja1.Cells(I + 1, 6) = Format(0, "#,##0.00")
                End If
            End If
            
            xlHoja1.Cells(I + 1, 7) = Format(xlHoja1.Cells(I + 1, 4) + xlHoja1.Cells(I + 1, 5) - xlHoja1.Cells(I + 1, 6), "#,##0.00")
            
            xlHoja1.Cells(I + 1, 8) = Format(IIf(Year(prRs!F_Compra) <> Me.mskAnio.Text, prRs!Dep_H_Ejer_Ant, 0), "#,##0.00")
            xlHoja1.Cells(I + 1, 9) = IIf(prRs!Baja = "False", Format(prRs!Dep_H_Mes_Ant, "#,##0.00"), Format(0, "#,##0.00"))
            xlHoja1.Cells(I + 1, 10) = Format(prRs!Dep_H_Mes, "#,##0.00")
            xlHoja1.Cells(I + 1, 11) = Format(xlHoja1.Cells(I + 1, 9) + xlHoja1.Cells(I + 1, 10), "#,##0.00")
            If prRs!Baja = "False" Then
                xlHoja1.Cells(I + 1, 12) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(I + 1, 12) = Format(prRs!Dep_H_Mes_Ant, "#,##0.00")
                Else
                    xlHoja1.Cells(I + 1, 12) = Format(0, "#,##0.00")
                End If
            End If
            xlHoja1.Cells(I + 1, 13) = Format(xlHoja1.Cells(I + 1, 8) + xlHoja1.Cells(I + 1, 11) - xlHoja1.Cells(I + 1, 12), "#,##0.00")
            xlHoja1.Cells(I + 1, 14) = Format(xlHoja1.Cells(I + 1, 13) - xlHoja1.Cells(I + 1, 7), "#,##0.00")
        End If
        prRs.MoveNext
    Wend
        
    xlHoja1.Cells.Select
    xlHoja1.Columns.AutoFit
    
    xlHoja1.Range("A1:N3").Select
    xlHoja1.Range("A1:N3").Font.Bold = True

    With xlHoja1.Range("A2:G2").Interior
        .ColorIndex = 36
        .Pattern = xlSolid
    End With
    With xlHoja1.Range("H2:N2").Interior
        .ColorIndex = 35
        .Pattern = xlSolid
    End With
    
    xlHoja1.Range("A2:N" & Trim(Str(I + 1))).Select
    xlHoja1.Range("A2:N" & Trim(Str(I + 1))).Borders(xlDiagonalDown).LineStyle = xlNone
    xlHoja1.Range("A2:N" & Trim(Str(I + 1))).Borders(xlDiagonalUp).LineStyle = xlNone
    
    With xlHoja1.Range("A2:N" & Trim(Str(I + 1))).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:N" & Trim(Str(I + 1))).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:N" & Trim(Str(I + 1))).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:N" & Trim(Str(I + 1))).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:N" & Trim(Str(I + 1))).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:N" & Trim(Str(I + 1))).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
End Sub

Private Sub GeneraReporteSaldoAjustado(prRs As ADODB.Recordset)
    Dim I As Integer
    Dim K As Integer
    Dim j As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lsSuma As String
    Dim sTipoGara As String
    Dim sTipoCred As String
    Dim lnAcum As Currency
    
    I = 1
    
    With xlHoja1.Range("A1:S1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlHoja1.Range("A1:T1").Merge
    xlHoja1.Range("A1:T1").FormulaR1C1 = " REPORTE DE SALDOS AJUSTADOS "
    
    With xlHoja1.Range("A2:L2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlHoja1.Range("A2:L2").Merge
    xlHoja1.Range("A2:L2").FormulaR1C1 = " COSTO DEL AJUSTADO DE ACTIVO FIJO "
    
    With xlHoja1.Range("M2:T2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlHoja1.Range("M2:T2").Merge
    xlHoja1.Range("M2:T2").FormulaR1C1 = " DEPRECIACION AJUSTADA DE ACTIVOS FIJOS "
    
    prRs.MoveFirst
    While Not prRs.EOF
        I = I + 1
        
        If I = 2 Then
            xlHoja1.Cells(I + 1, 1) = "COD CTIVO"
            xlHoja1.Cells(I + 1, 2) = "COD. PATRIM."
            xlHoja1.Cells(I + 1, 3) = "F. ADQUIS"
            xlHoja1.Cells(I + 1, 4) = "SALDO ACT HIST"
            xlHoja1.Cells(I + 1, 5) = "SALDO AJUS AÑO ANT"
            xlHoja1.Cells(I + 1, 6) = "COM DEL AÑO"
            xlHoja1.Cells(I + 1, 7) = "RET DEL AÑO"
            xlHoja1.Cells(I + 1, 8) = "FAC DE AJUS"
            xlHoja1.Cells(I + 1, 9) = "REEXP VAL AJUS ANT"
            xlHoja1.Cells(I + 1, 10) = "COM AJUS"
            xlHoja1.Cells(I + 1, 11) = "RET AJUS"
            xlHoja1.Cells(I + 1, 12) = "VAL ACT AJUS"
            xlHoja1.Cells(I + 1, 13) = "DEP ACUM AÑO ANT"
            xlHoja1.Cells(I + 1, 14) = "REEXP DEP AJUS AÑO ANT"
            xlHoja1.Cells(I + 1, 15) = "DEP AJUS EJER MES ANT"
            xlHoja1.Cells(I + 1, 16) = "DEP AJUST DEL MES"
            xlHoja1.Cells(I + 1, 17) = "TOT DEP DEL EJER AJSU"
            xlHoja1.Cells(I + 1, 18) = "DEP AJUS ACUM RET"
            xlHoja1.Cells(I + 1, 19) = "DEP AJUS ACUM TOTAL"
            xlHoja1.Cells(I + 1, 20) = "VALOR EN LIBROS AJUS"
            
            I = I + 1

            xlHoja1.Cells(I + 1, 1) = prRs!Codigo & "-" & prRs!Serie
            xlHoja1.Cells(I + 1, 2) = Format(I - 2, "00000")
            xlHoja1.Cells(I + 1, 3) = "'" & Format(prRs!F_Compra, gsFormatoFechaView)
            xlHoja1.Cells(I + 1, 4) = Format(prRs!Valor, "#,##0.00")
            xlHoja1.Cells(I + 1, 5) = IIf(Year(prRs!F_Compra) = Me.mskAnio, 0, Format(prRs!Valor_Ajustado, "#,##0.00"))
            xlHoja1.Cells(I + 1, 6) = Format(IIf(Year(prRs!F_Compra) = Me.mskAnio, prRs!Valor, 0), "#,##0.00")
            
            If prRs!Baja = "False" Or Not IsNumeric(prRs!F_Baja) Then
                xlHoja1.Cells(I + 1, 7) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(I + 1, 7) = Format(prRs!Valor, "#,##0.00")
                Else
                    xlHoja1.Cells(I + 1, 7) = Format(0, "#,##0.00")
                End If
            End If
            
            xlHoja1.Cells(I + 1, 8) = Format(prRs!F_Ajuste, "#,##0.000")
            xlHoja1.Cells(I + 1, 9) = Format(IIf(Year(prRs!F_Compra) = Me.mskAnio, 0, prRs!Valor_Ajustado) * prRs!F_Ajuste, "#,##0.00")
            xlHoja1.Cells(I + 1, 10) = Format(IIf(Year(prRs!F_Compra) = Year(gdFecSis), prRs!Valor, 0) * prRs!F_Ajuste, "#,##0.00")
            
            If prRs!Baja = "False" Or Not IsNumeric(prRs!F_Baja) Then
                xlHoja1.Cells(I + 1, 11) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(I + 1, 11) = Format(prRs!Valor * prRs!F_Ajuste, "#,##0.00")
                Else
                    xlHoja1.Cells(I + 1, 11) = Format(0, "#,##0.00")
                End If
            End If
            
            xlHoja1.Cells(I + 1, 12) = Format(xlHoja1.Cells(I + 1, 9) + xlHoja1.Cells(I + 1, 10) - xlHoja1.Cells(I + 1, 11), "#,##0.00")
            
            xlHoja1.Cells(I + 1, 13) = Format(IIf(Year(prRs!F_Compra) >= Me.mskAnio, 0, prRs!Valor_Ajustado), "#,##0.00")
            xlHoja1.Cells(I + 1, 14) = Format(IIf(Year(prRs!F_Compra) >= Me.mskAnio, 0, prRs!Valor_Ajustado) * prRs!F_Ajuste, "#,##0.00")
            xlHoja1.Cells(I + 1, 15) = IIf(prRs!Baja = "False", Format(prRs!Dep_A_Mes_Ant, "#,##0.00"), Format(0, "#,##0.00"))
            xlHoja1.Cells(I + 1, 16) = Format(prRs!Dep_A_Mes, "#,##0.00")
            xlHoja1.Cells(I + 1, 17) = Format(xlHoja1.Cells(I + 1, 15) + xlHoja1.Cells(I + 1, 16), "#,##0.00")
            xlHoja1.Cells(I + 1, 18) = IIf(prRs!Baja <> "False", Format(prRs!Dep_A_Mes_Ant, "#,##0.00"), Format(0, "#,##0.00"))
            xlHoja1.Cells(I + 1, 19) = Format(xlHoja1.Cells(I + 1, 14) + xlHoja1.Cells(I + 1, 17) - xlHoja1.Cells(I + 1, 18), "#,##0.00")
            xlHoja1.Cells(I + 1, 20) = Format(xlHoja1.Cells(I + 1, 12) - xlHoja1.Cells(I + 1, 19), "#,##0.00")
        Else
            xlHoja1.Cells(I + 1, 1) = prRs!Codigo & "-" & prRs!Serie
            xlHoja1.Cells(I + 1, 2) = Format(I - 2, "00000")
            xlHoja1.Cells(I + 1, 3) = "'" & Format(prRs!F_Compra, gsFormatoFechaView)
            xlHoja1.Cells(I + 1, 4) = Format(prRs!Valor, "#,##0.00")
            xlHoja1.Cells(I + 1, 5) = IIf(Year(prRs!F_Compra) = Me.mskAnio, 0, Format(prRs!Valor_Ajustado, "#,##0.00"))
            xlHoja1.Cells(I + 1, 6) = Format(IIf(Year(prRs!F_Compra) = Year(gdFecSis), prRs!Valor, 0), "#,##0.00")
            
            If prRs!Baja = "False" Or Not IsNumeric(prRs!F_Baja) Then
                xlHoja1.Cells(I + 1, 7) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(I + 1, 7) = Format(prRs!Valor, "#,##0.00")
                Else
                    xlHoja1.Cells(I + 1, 7) = Format(0, "#,##0.00")
                End If
            End If
            
            xlHoja1.Cells(I + 1, 8) = Format(prRs!F_Ajuste, "#,##0.000")
            xlHoja1.Cells(I + 1, 9) = Format(IIf(Year(prRs!F_Compra) = Me.mskAnio, 0, prRs!Valor_Ajustado) * prRs!F_Ajuste, "#,##0.00")
            xlHoja1.Cells(I + 1, 10) = Format(IIf(Year(prRs!F_Compra) = Year(gdFecSis), prRs!Valor, 0) * prRs!F_Ajuste, "#,##0.00")
            
            If prRs!Baja = "False" Or Not IsNumeric(prRs!F_Baja) Then
                xlHoja1.Cells(I + 1, 11) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(I + 1, 11) = Format(prRs!Valor * prRs!F_Ajuste, "#,##0.00")
                Else
                    xlHoja1.Cells(I + 1, 11) = Format(0, "#,##0.00")
                End If
            End If
            
            xlHoja1.Cells(I + 1, 12) = Format(xlHoja1.Cells(I + 1, 9) + xlHoja1.Cells(I + 1, 10) - xlHoja1.Cells(I + 1, 11), "#,##0.00")
            
            xlHoja1.Cells(I + 1, 13) = Format(IIf(Year(prRs!F_Compra) >= Me.mskAnio, 0, prRs!Valor_Ajustado), "#,##0.00")
            xlHoja1.Cells(I + 1, 14) = Format(IIf(Year(prRs!F_Compra) >= Me.mskAnio, 0, prRs!Valor_Ajustado) * prRs!F_Ajuste, "#,##0.00")
            xlHoja1.Cells(I + 1, 15) = IIf(prRs!Baja = "False", Format(prRs!Dep_A_Mes_Ant, "#,##0.00"), Format(0, "#,##0.00"))
            xlHoja1.Cells(I + 1, 16) = Format(prRs!Dep_A_Mes, "#,##0.00")
            xlHoja1.Cells(I + 1, 17) = Format(xlHoja1.Cells(I + 1, 15) + xlHoja1.Cells(I + 1, 16), "#,##0.00")
            xlHoja1.Cells(I + 1, 18) = IIf(prRs!Baja <> "False", Format(prRs!Dep_A_Mes_Ant, "#,##0.00"), Format(0, "#,##0.00"))
            xlHoja1.Cells(I + 1, 19) = Format(xlHoja1.Cells(I + 1, 14) + xlHoja1.Cells(I + 1, 17) - xlHoja1.Cells(I + 1, 18), "#,##0.00")
            xlHoja1.Cells(I + 1, 20) = Format(-xlHoja1.Cells(I + 1, 12) + xlHoja1.Cells(I + 1, 19), "#,##0.00")

        End If
        prRs.MoveNext
    Wend
        
    xlHoja1.Select
    xlHoja1.Columns.AutoFit
    
    xlHoja1.Range("A1:T3").Select
    xlHoja1.Range("A1:T3").Font.Bold = True

    With xlHoja1.Range("A2:L2").Interior
        .ColorIndex = 36
        .Pattern = xlSolid
    End With
    With xlHoja1.Range("M2:T2").Interior
        .ColorIndex = 35
        .Pattern = xlSolid
    End With
    
    xlHoja1.Range("A2:T" & Trim(Str(I + 1))).Select
    xlHoja1.Range("A2:T" & Trim(Str(I + 1))).Borders(xlDiagonalDown).LineStyle = xlNone
    xlHoja1.Range("A2:T" & Trim(Str(I + 1))).Borders(xlDiagonalUp).LineStyle = xlNone
    
    With xlHoja1.Range("A2:T" & Trim(Str(I + 1))).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:T" & Trim(Str(I + 1))).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:T" & Trim(Str(I + 1))).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:T" & Trim(Str(I + 1))).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:T" & Trim(Str(I + 1))).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:T" & Trim(Str(I + 1))).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End Sub

Private Sub llenagrid()
    Dim oDepo As DAgencia
    Set oDepo = New DAgencia
    Dim ldFecha As Date, I As Integer
    Dim rs1 As ADODB.Recordset
    Dim rs9 As ADODB.Recordset
    
    Set rs9 = New ADODB.Recordset
    
    FlexEdit1.Clear
    ldFecha = CDate("01/" & Format(Trim(Right(Me.cmbMes.Text, 5)), "00") & "/" & Me.mskAnio.Text)
    
    FlexEdit1.rsFlex = oDepo.ObtieneResumenCredPigno(Format(ldFecha, "yyyymmdd"), Format(Trim(Right(Me.cboTpo.Text, 3)), "00"))
    
End Sub
