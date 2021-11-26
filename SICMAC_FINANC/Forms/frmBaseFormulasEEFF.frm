VERSION 5.00
Begin VB.Form frmNIIFBaseFormulasEEFF 
   Caption         =   "Estados Financiero en Base a Fórmulas"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11025
   Icon            =   "frmBaseFormulasEEFF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdProyecciones 
      Caption         =   "&Proyecciones"
      Height          =   375
      Left            =   360
      TabIndex        =   14
      ToolTipText     =   "Proyecciones"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   9360
      TabIndex        =   13
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9480
      TabIndex        =   9
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdReplicar 
      Caption         =   "Replicar"
      Height          =   375
      Left            =   9480
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdBajar 
      Caption         =   "Bajar"
      Height          =   375
      Left            =   9480
      TabIndex        =   7
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdSubir 
      Caption         =   "Subir"
      Height          =   375
      Left            =   9480
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "Quitar"
      Height          =   375
      Left            =   9480
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   9480
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin Sicmact.FlexEdit feNotas 
      Height          =   3255
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5741
      Cols0           =   7
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "-Nivel-Concepto-Fórmula consolidado-Fórmula <= 2012-Tipo-nCorreInt"
      EncabezadosAnchos=   "0-1200-2800-1700-1600-1200-0"
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
      ColumnasAEditar =   "X-1-2-3-4-5-X"
      ListaControles  =   "0-3-0-0-0-3-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-L-L-L-C"
      FormatosEdit    =   "0-0-0-0-0-0-0"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.ComboBox cboTipoReporte 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   5415
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   240
      TabIndex        =   12
      Top             =   4320
      Width           =   10695
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Height          =   3255
      Left            =   9360
      TabIndex        =   11
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Registros del reporte"
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
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Selecione tipo de reporte"
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
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmNIIFBaseFormulasEEFF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fsOpeCod As String
Dim fsOpeDesc As String
Dim fMatNotasDet As Variant
Dim MatNotasEstadoDetInicio As Variant
Dim fbAceptar As Boolean
Private Type TCtaCont
    CuentaContable As String
    Saldo As Currency
End Type
Private Sub cmdExportar_Click()
    Call frmNIIFBaseFormulasEEFFRep.Inicio(fsOpeCod, fsOpeDesc)
End Sub

Private Sub cmdGrabar_Click()
    Dim oRep As New DRepFormula
    Dim lsMovNro As String
    Dim lbExito As Boolean
    Dim MatNotas As Variant
    Dim i As Long
    Dim bTrans As Boolean
    
    'If validarGrabar = False Then Exit Sub
    If validarRegistroDatosNotasEstado = False Then Exit Sub
    
    If MsgBox("¿Esta seguro de guardar la configuración de las Notas de Estado?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo ErrGrabar
    Screen.MousePointer = 11
    oRep.dBeginTrans
    bTrans = True
    
    oRep.GuardarProyeccionTemporal fsOpeCod
    Call oRep.EliminarEEFF(fsOpeCod)
    
    ReDim MatNotas(1 To 6, 0)
    For i = 1 To feNotas.Rows - 1
        ReDim Preserve MatNotas(1 To 6, 1 To i)
        MatNotas(1, i) = Trim(Right(Trim(feNotas.TextMatrix(i, 1)), 5)) 'Nivel
        MatNotas(2, i) = Trim(feNotas.TextMatrix(i, 2)) 'descripcion
        MatNotas(3, i) = Trim(feNotas.TextMatrix(i, 3)) 'FormulaCons
        MatNotas(4, i) = Trim(feNotas.TextMatrix(i, 4)) 'FormulaAgencia
        MatNotas(5, i) = IIf(MatNotas(5, i) = "", "1", Trim(Right(Trim(feNotas.TextMatrix(i, 5)), 5))) 'Tipo
        MatNotas(6, i) = feNotas.TextMatrix(i, 6) 'Correlativo
    Call oRep.InsertarEEFF(fsOpeCod, i, MatNotas(1, i), MatNotas(2, i), MatNotas(3, i), MatNotas(4, i), MatNotas(5, i), gsCodUser)
    If MatNotas(6, i) <> "" Then
        oRep.MigrarProyeccion fsOpeCod, MatNotas(6, i), i
    End If
    Next
    
    oRep.dCommitTrans
    bTrans = False
    
    For i = 1 To feNotas.Rows - 1 'Reiniciamos los correlativos de las Notas
        feNotas.TextMatrix(i, 6) = feNotas.TextMatrix(i, 0)
    Next
    
    Screen.MousePointer = 0
    'If feNotas.Rows > 1 Then
        MsgBox "Se ha grabado satisfactoriamente los cambios de las Notas Estado", vbInformation, "Aviso"
    'End If

    Set oRep = Nothing
    Set MatNotas = Nothing
    Exit Sub
ErrGrabar:
    Screen.MousePointer = 0
    If bTrans Then
        oRep.dRollbackTrans
        Set oRep = Nothing
    End If
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdProyecciones_Click()
    frmNIIFBaseFormulasEEFFProy.Inicio fsOpeCod
End Sub

Private Sub cmdQuitar_Click()
    feNotas.EliminaFila feNotas.row
End Sub

Private Sub cmdReplicar_Click()
    feNotas.TextMatrix(feNotas.row, 4) = ObtenerResultadoFormula(feNotas.TextMatrix(feNotas.row, 3))
End Sub
Private Function ObtenerResultadoFormula(ByVal psFormula As String) As String
    Dim oBal As New DbalanceCont
    Dim oFormula As New NInterpreteFormula
    Dim lsFormula As String, lsTmp As String, lsTmp1 As String, lsCadFormula As String
    Dim MatDatos() As TCtaCont
    Dim i As Long, j As Long, nCtaCont As Long
    
    lsFormula = Trim(psFormula)
    ReDim MatDatos(0)
    nCtaCont = 0

    For i = 1 To Len(lsFormula)
        If ((Mid(Trim(lsFormula), i, 1) >= "0" And Mid(Trim(lsFormula), i, 1) <= "9") Or UCase(Mid(Trim(lsFormula), i, 1)) = "M") Then
            lsTmp = lsTmp + UCase(Mid(Trim(lsFormula), i, 1))
        Else
            If Len(lsTmp) > 0 Then
                nCtaCont = nCtaCont + 1
                ReDim Preserve MatDatos(nCtaCont)
                MatDatos(nCtaCont).CuentaContable = lsTmp + "AG"
            End If
            lsTmp = ""
        End If
    Next i
    If Len(lsTmp) > 0 Then
        nCtaCont = nCtaCont + 1
        ReDim Preserve MatDatos(nCtaCont)
        MatDatos(nCtaCont).CuentaContable = lsTmp + "AG"
    End If
    'Genero la formula en cadena
    lsTmp = ""
    lsCadFormula = ""
    For i = 1 To Len(lsFormula)
        If (Mid(Trim(lsFormula), i, 1) >= "0" And Mid(Trim(lsFormula), i, 1) <= "9") Or UCase(Mid(Trim(lsFormula), i, 1)) = "M" Then
            lsTmp = lsTmp + UCase(Mid(Trim(lsFormula), i, 1))
        Else
            If Len(lsTmp) > 0 Then
                For j = 1 To nCtaCont
                    If MatDatos(j).CuentaContable = lsTmp + "AG" Then
                        lsCadFormula = lsCadFormula & MatDatos(j).CuentaContable
                        Exit For
                    End If
                Next j
            End If
            lsTmp = ""
            lsCadFormula = lsCadFormula & UCase(Mid(Trim(lsFormula), i, 1))
        End If
    Next
    If Len(lsTmp) > 0 Then
        For j = 1 To nCtaCont
           If MatDatos(j).CuentaContable = lsTmp + "AG" Then
               lsCadFormula = lsCadFormula & MatDatos(j).CuentaContable
               Exit For
           End If
        Next j
    End If
    ObtenerResultadoFormula = lsCadFormula
    Set oBal = Nothing
    Set oFormula = Nothing
End Function

Private Sub cmdSalir_Click()
    fbAceptar = False
    Hide
End Sub

Private Sub Form_Load()
    CentraForm Me
    IniciarControles
    Call llenarComboTipoReporte
End Sub
Private Sub llenarComboTipoReporte()
    cboTipoReporte.AddItem fsOpeCod & "-" & fsOpeDesc & Space(200) & fsOpeCod
    cboTipoReporte.ListIndex = IndiceListaCombo(cboTipoReporte, fsOpeCod)
    If Trim(cboTipoReporte.Text) <> "" Then
        cmdAgregar.Enabled = True
    End If
End Sub
Private Sub ListarConfiguracionEEFF(ByVal psOpeCod As String)
    Dim oRep As New DRepFormula
    Dim rsNotas As New ADODB.Recordset
    Dim rsNotasDet As New ADODB.Recordset
    Dim Detalle As Variant
    Dim iCab As Long, iDet As Long
    
    Set rsNotas = oRep.ObtenerEEFF(psOpeCod)
    Call LimpiaFlex(feNotas)

    Set fMatNotasDet = Nothing
    ReDim fMatNotasDet(0)

    If Not RSVacio(rsNotas) Then
        ReDim fMatNotasDet(1 To rsNotas.RecordCount)
        For iCab = 1 To rsNotas.RecordCount
            feNotas.AdicionaFila
            'Notas
            feNotas.TextMatrix(feNotas.row, 1) = IIf(rsNotas!nNivelCod = "1", "TITULO                   1", IIf(rsNotas!nNivelCod = "2", "SUBTITULO                     2", "CONCEPTO                       3"))
            feNotas.TextMatrix(feNotas.row, 2) = rsNotas!cConceptoDesc
            feNotas.TextMatrix(feNotas.row, 3) = rsNotas!cFormulaCons
            feNotas.TextMatrix(feNotas.row, 4) = rsNotas!cFormulaAgen
            feNotas.TextMatrix(feNotas.row, 5) = IIf(rsNotas!nTipoCod = "1", "NEUTRO                    1", IIf(rsNotas!nTipoCod = "2", "DEUDOR                         2", "ACREEDOR                       3"))
            feNotas.TextMatrix(feNotas.row, 6) = rsNotas!nCorreInt
            rsNotas.MoveNext
        Next
            cmdAgregar.Enabled = True
            cmdSubir.Enabled = True
            cmdBajar.Enabled = True
            cmdQuitar.Enabled = True
            cmdReplicar.Enabled = True
            cmdExportar.Enabled = True
            'cmdAgregarP.Enabled = True
        feNotas.TopRow = 1
        feNotas.row = 1
    End If
    Set oRep = Nothing
    Set rsNotas = Nothing
    Set rsNotasDet = Nothing
    Set Detalle = Nothing
End Sub
Private Sub feNotas_RowColChange()
    Dim rsOpt As New ADODB.Recordset
    Dim i As Integer
    
    If feNotas.Col = 1 Or feNotas.Col = 5 Then
        With rsOpt
            .Fields.Append "desc", adVarChar, 10
            .Fields.Append "value", adVarChar, 2
        End With
        If feNotas.Col = 1 Then
            With rsOpt
                .Open
                .AddNew
                .Fields("desc") = "TITULO"
                .Fields("value") = "1"
                .AddNew
                .Fields("desc") = "SUBTITULO"
                .Fields("value") = "2"
                .AddNew
                .Fields("desc") = "CONCEPTO"
                .Fields("value") = "3"
            End With
        ElseIf feNotas.Col = 5 Then
            With rsOpt
                .Open
                .AddNew
                .Fields("desc") = "NEUTRO"
                .Fields("value") = "1"
                .AddNew
                .Fields("desc") = "ACREEDOR"
                .Fields("value") = "2"
                .AddNew
                .Fields("desc") = "DEUDOR"
                .Fields("value") = "3"
            End With
        End If
        rsOpt.MoveFirst
        feNotas.CargaCombo rsOpt
    End If
    Set rsOpt = Nothing
End Sub
Private Sub IniciarControles()
    Call ListarConfiguracionEEFF(fsOpeCod)
    cmdAgregar.Enabled = False
    cmdSubir.Enabled = False
    cmdBajar.Enabled = False
    cmdQuitar.Enabled = False
    cmdReplicar.Enabled = False
    cmdExportar.Enabled = False
    'cmdAgregarP.Enabled = False
    If fsOpeCod = gContRepEstadoSitFinanEEFF1 Or fsOpeCod = gContRepEstadoSitFinanEEFF2 Then 'EJVG20140909
        cmdProyecciones.Visible = True
    End If
End Sub
Public Sub Inicio(ByVal psOpeCod As String, psOpeDesc As String)
    fsOpeCod = psOpeCod
    fsOpeDesc = psOpeDesc
    Caption = "CONFIGURACIÓN " & UCase(psOpeDesc)
    Call ListarConfiguracionEEFF(fsOpeCod)
    Show 1
End Sub
Private Function validarRegistroDatosEEFF() As Boolean
    validarRegistroDatosEEFF = True
    Dim i As Long, j As Long
    For i = 1 To feNotas.Rows - 1 'valida fila x fila
        For j = 1 To feNotas.Cols - 2 '2 xq el ultimo es aux
            If j = 1 Or j = 3 Or j = 4 Then 'xq las formulas y la desc son opcionales
                If Trim(feNotas.TextMatrix(i, j)) = "" Then
                    validarRegistroDatosEEFF = False
                    MsgBox "Ud. debe de ingresar el dato '" & UCase(feNotas.TextMatrix(0, j)) & "'", vbInformation, "Aviso"
                    feNotas.row = i
                    feNotas.Col = j
                    feNotas.SetFocus
                    Exit Function
                End If
            End If
        Next
    Next
End Function
Private Sub cmdSubir_Click()
    Dim lsTipoDetalle1 As String, lsTipoDetalle2 As String
    Dim lsDesc1 As String, lsDesc2 As String
    Dim lsNivel1 As String, lsNivel2 As String
    Dim lsNegrita1 As String, lsNegrita2 As String
    Dim lsFormula1_1 As String, lsFormula1_2 As String
    Dim lsFormula1_2012_1 As String, lsFormula1_2012_2 As String
    Dim lsFormula2_1 As String, lsFormula2_2 As String
    Dim lsFormula2_2012_1 As String, lsFormula2_2012_2 As String
    Dim lsFormula3_1 As String, lsFormula3_2 As String
    Dim lsFormula3_2012_1 As String, lsFormula3_2012_2 As String
    Dim lsFormula4_1 As String, lsFormula4_2 As String
    Dim lsFormula4_2012_1 As String, lsFormula4_2012_2 As String
    Dim lsFormula5_1 As String, lsFormula5_2 As String
    Dim lsFormula5_2012_1 As String, lsFormula5_2012_2 As String
    Dim lsCorreInt1 As String, lsCorreInt2 As String 'EJVG20140915
    
    If validarRegistroDatosEEFF = False Then Exit Sub

    If feNotas.row > 1 Then
        lsTipoDetalle1 = feNotas.TextMatrix(feNotas.row - 1, 1)
        lsDesc1 = feNotas.TextMatrix(feNotas.row - 1, 2)
        lsNivel1 = feNotas.TextMatrix(feNotas.row - 1, 3)
        lsNegrita1 = feNotas.TextMatrix(feNotas.row - 1, 4)
        lsFormula1_1 = feNotas.TextMatrix(feNotas.row - 1, 5)
        lsCorreInt1 = feNotas.TextMatrix(feNotas.row - 1, 6)
        
        lsTipoDetalle2 = feNotas.TextMatrix(feNotas.row, 1)
        lsDesc2 = feNotas.TextMatrix(feNotas.row, 2)
        lsNivel2 = feNotas.TextMatrix(feNotas.row, 3)
        lsNegrita2 = feNotas.TextMatrix(feNotas.row, 4)
        lsFormula1_2 = feNotas.TextMatrix(feNotas.row, 5)
        lsCorreInt2 = feNotas.TextMatrix(feNotas.row, 6)
        
        feNotas.TextMatrix(feNotas.row - 1, 1) = lsTipoDetalle2
        feNotas.TextMatrix(feNotas.row - 1, 2) = lsDesc2
        feNotas.TextMatrix(feNotas.row - 1, 3) = lsNivel2
        feNotas.TextMatrix(feNotas.row - 1, 4) = lsNegrita2
        feNotas.TextMatrix(feNotas.row - 1, 5) = lsFormula1_2
        feNotas.TextMatrix(feNotas.row - 1, 6) = lsCorreInt2
        
        feNotas.TextMatrix(feNotas.row, 1) = lsTipoDetalle1
        feNotas.TextMatrix(feNotas.row, 2) = lsDesc1
        feNotas.TextMatrix(feNotas.row, 3) = lsNivel1
        feNotas.TextMatrix(feNotas.row, 4) = lsNegrita1
        feNotas.TextMatrix(feNotas.row, 5) = lsFormula1_1
        feNotas.TextMatrix(feNotas.row, 6) = lsCorreInt1
               
        feNotas.row = feNotas.row - 1
        feNotas.SetFocus
    End If
End Sub
Private Sub cmdBajar_Click()
    Dim lsTipoDetalle1 As String, lsTipoDetalle2 As String
    Dim lsDesc1 As String, lsDesc2 As String
    Dim lsNivel1 As String, lsNivel2 As String
    Dim lsNegrita1 As String, lsNegrita2 As String
    Dim lsFormula1_1 As String, lsFormula1_2 As String
    Dim lsFormula1_2012_1 As String, lsFormula1_2012_2 As String
    Dim lsFormula2_1 As String, lsFormula2_2 As String
    Dim lsFormula2_2012_1 As String, lsFormula2_2012_2 As String
    Dim lsFormula3_1 As String, lsFormula3_2 As String
    Dim lsFormula3_2012_1 As String, lsFormula3_2012_2 As String
    Dim lsFormula4_1 As String, lsFormula4_2 As String
    Dim lsFormula4_2012_1 As String, lsFormula4_2012_2 As String
    Dim lsFormula5_1 As String, lsFormula5_2 As String
    Dim lsFormula5_2012_1 As String, lsFormula5_2012_2 As String
    Dim lsCorreInt1 As String, lsCorreInt2 As String 'EJVG20140915

    If validarRegistroDatosEEFF = False Then Exit Sub

    If feNotas.row < feNotas.Rows - 1 Then
        lsTipoDetalle1 = feNotas.TextMatrix(feNotas.row + 1, 1)
        lsDesc1 = feNotas.TextMatrix(feNotas.row + 1, 2)
        lsNivel1 = feNotas.TextMatrix(feNotas.row + 1, 3)
        lsNegrita1 = feNotas.TextMatrix(feNotas.row + 1, 4)
        lsFormula1_1 = feNotas.TextMatrix(feNotas.row + 1, 5)
        lsCorreInt1 = feNotas.TextMatrix(feNotas.row + 1, 6)
        
        lsTipoDetalle2 = feNotas.TextMatrix(feNotas.row, 1)
        lsDesc2 = feNotas.TextMatrix(feNotas.row, 2)
        lsNivel2 = feNotas.TextMatrix(feNotas.row, 3)
        lsNegrita2 = feNotas.TextMatrix(feNotas.row, 4)
        lsFormula1_2 = feNotas.TextMatrix(feNotas.row, 5)
        lsCorreInt2 = feNotas.TextMatrix(feNotas.row, 6)
        
        feNotas.TextMatrix(feNotas.row + 1, 1) = lsTipoDetalle2
        feNotas.TextMatrix(feNotas.row + 1, 2) = lsDesc2
        feNotas.TextMatrix(feNotas.row + 1, 3) = lsNivel2
        feNotas.TextMatrix(feNotas.row + 1, 4) = lsNegrita2
        feNotas.TextMatrix(feNotas.row + 1, 5) = lsFormula1_2
        feNotas.TextMatrix(feNotas.row + 1, 6) = lsCorreInt2
        
        feNotas.TextMatrix(feNotas.row, 1) = lsTipoDetalle1
        feNotas.TextMatrix(feNotas.row, 2) = lsDesc1
        feNotas.TextMatrix(feNotas.row, 3) = lsNivel1
        feNotas.TextMatrix(feNotas.row, 4) = lsNegrita1
        feNotas.TextMatrix(feNotas.row, 5) = lsFormula1_1
        feNotas.TextMatrix(feNotas.row, 6) = lsCorreInt1
        feNotas.row = feNotas.row + 1
        feNotas.SetFocus
    End If
End Sub

Private Sub RecuperaNotasEstadoDetalle()
    Dim Mat As Variant
    Dim i As Long
    
    'ReDim mat(1 To 5, 0)
    ReDim Mat(1 To 6, 0) 'EJVG20140915
    If fbAceptar Then
        If Not (feNotas.Rows - 1 = 1 And Len(Trim(feNotas.TextMatrix(1, 0))) = 0) Then 'Flex no esta Vacio
            For i = 1 To feNotas.Rows - 1
                ReDim Preserve Mat(1 To 6, 1 To i)
                Mat(1, i) = Trim(Right(feNotas.TextMatrix(i, 1), 2)) 'Tipo de Detalle
                Mat(2, i) = feNotas.TextMatrix(i, 2) 'Descripcion
                Mat(3, i) = Trim(Right(feNotas.TextMatrix(i, 3), 2)) 'Nivel
                Mat(4, i) = Trim(Right(feNotas.TextMatrix(i, 4), 2)) 'Negrita
                Mat(5, i) = feNotas.TextMatrix(i, 5) 'Formula 1
                Mat(6, i) = feNotas.TextMatrix(i, 6) 'Correlativo
            Next
            cmdAgregar.Enabled = True
            cmdSubir.Enabled = True
            cmdBajar.Enabled = True
            cmdQuitar.Enabled = True
            cmdReplicar.Enabled = True
            cmdExportar.Enabled = True
            cmdAgregar.Enabled = True
        End If
    End If
    Set Mat = Nothing
End Sub

Private Sub feNotas_Click()
    If feNotas.row > 0 Then
        If feNotas.TextMatrix(feNotas.row, 0) <> "" Then
            cmdSubir.Enabled = True
            cmdBajar.Enabled = True
        End If
    End If
End Sub
Private Function validarRegistroDatosNotasEstado() As Boolean
    validarRegistroDatosNotasEstado = True
    Dim i As Long, j As Long
    For i = 1 To feNotas.Rows - 1 'valida fila x fila
        For j = 1 To feNotas.Cols - 2 '2 xq el ultimo es aux
            If j <> 2 Then 'xq la plantilla contable es opcional
                If Trim(feNotas.TextMatrix(i, j)) = "" Then
'                    validarRegistroDatosNotasEstado = False
'                    MsgBox "Ud. debe de ingresar el dato '" & UCase(feNotas.TextMatrix(0, j)) & "'", vbInformation, "Aviso"
'                    feNotas.Row = I
'                    feNotas.Col = j
'                    feNotas.SetFocus
'                    Exit Function
                End If
            End If
        Next
    Next
End Function
Private Sub cmdAgregar_Click()
    Dim MatDetalle As Variant
    ReDim MatDetalle(1 To 14, 0)
    
    If Not (feNotas.Rows - 1 = 1 And Len(Trim(feNotas.TextMatrix(1, 0))) = 0) Then 'Flex no esta Vacio
        If validarRegistroDatosNotasEstado = False Then Exit Sub
    End If
    
    feNotas.AdicionaFila
    ReDim Preserve fMatNotasDet(1 To UBound(fMatNotasDet) + 1)
    fMatNotasDet(UBound(fMatNotasDet)) = MatDetalle

    feNotas.SetFocus
    SendKeys "{Enter}"
    cmdQuitar.Enabled = True
    cmdReplicar.Enabled = True
    Call feNotas_RowColChange
End Sub
