VERSION 5.00
Begin VB.Form frmIntangibleAmortizacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Intangibles - Amortización"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16215
   Icon            =   "frmIntangibleAmortizacion.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   16215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAmortizar 
      Caption         =   "Amortizar"
      Height          =   375
      Left            =   13560
      TabIndex        =   9
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   14760
      TabIndex        =   8
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CheckBox chkEstadistico 
      Caption         =   "Estadístico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   4320
      Width           =   1335
   End
   Begin Sicmact.FlexEdit feAmortizar 
      Height          =   2775
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   15975
      _ExtentX        =   28178
      _ExtentY        =   4895
      Cols0           =   15
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Cod.-Descripción-Rubro-Moneda-Valor-Valor MN-Nº de Amort.-Monto Amortizar--nMovNro-cMovnro-CtaCont-Estado-FechaAct"
      EncabezadosAnchos=   "300-1500-6000-1500-1000-1200-1200-1200-1200-350-0-0-0-0-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-9-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-4-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "L-L-L-C-C-R-C-R-R-C-L-C-C-R-L"
      FormatosEdit    =   "0-1-1-1-1-2-2-3-2-0-1-1-1-3-1"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Frame frmBuscar 
      Caption         =   "Buscar"
      Height          =   975
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin VB.CommandButton cmdBuscar 
         Appearance      =   0  'Flat
         Caption         =   "Mostrar"
         Height          =   345
         Left            =   6720
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtAnio 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5640
         MaxLength       =   4
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox cboMes 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmIntangibleAmortizacion.frx":030A
         Left            =   3600
         List            =   "frmIntangibleAmortizacion.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox cboRubro 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Rubro:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmIntangibleAmortizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**-------------------------------------------------------------------------------------**'
'** Formulario : frmIntangibleAmortizacion
'** Finalidad  : Este formulario permite Amortizar las intangibles que
'**                han sido activadas
'** Programador: Paolo Hector Sinti Cabrera - PASI
'** Fecha/Hora : 20140305 11:50 AM
'**-------------------------------------------------------------------------------------**'

Option Explicit
Dim oIntangible As dIntangible
Dim lsErrores() As String
Dim ldFecIni As Date
Dim ldFecFin As Date
Dim lsRubroc As String
Dim lsCadenaAsiento As String
Dim lnNroAsiento() As String
Dim lsNroAsientoRef() As String

Private Sub cmdBuscar_Click()
    CargarDatos
End Sub
Private Sub CargarDatos()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oIntangible = New dIntangible

    Dim lsRubro As String
    Dim row As Integer
    
    lsRubro = IIf(Right(cboRubro.Text, 1) = 0, "%", Right(cboRubro.Text, 1))
    lsRubroc = Left(cboRubro.Text, 50)
    ldFecIni = CDate("01/" & Format(cboMes.ListIndex + 1, "00") & "/" & txtAnio.Text)
    ldFecFin = DateAdd("M", 1, ldFecIni) - 1
    
    FormateaFlex feAmortizar
    Set rs = oIntangible.ListaIntangibleAmortizar(lsRubro, ldFecIni, ldFecFin)
    If Not rs.EOF Then
        Do While Not rs.EOF
            feAmortizar.AdicionaFila
            row = feAmortizar.row
            feAmortizar.TextMatrix(row, 1) = rs!Codigo
            feAmortizar.TextMatrix(row, 2) = rs!Descripcion
            feAmortizar.TextMatrix(row, 3) = rs!Rubro
            feAmortizar.TextMatrix(row, 4) = rs!Moneda
            feAmortizar.TextMatrix(row, 5) = Format(rs!Valor, "#,#0.00")
            feAmortizar.TextMatrix(row, 6) = Format(rs!ValorMN, "#,#0.00")
            feAmortizar.TextMatrix(row, 7) = rs!NMesAmort
            feAmortizar.TextMatrix(row, 8) = Format(rs!MontoAmor, "#,#0.00")
            feAmortizar.TextMatrix(row, 9) = rs!Estado
            feAmortizar.TextMatrix(row, 10) = rs!nMovNro
            feAmortizar.TextMatrix(row, 11) = rs!cMovNro
            feAmortizar.TextMatrix(row, 12) = rs!CtaCont
            feAmortizar.TextMatrix(row, 13) = rs!Estado
            feAmortizar.TextMatrix(row, 14) = rs!FechaAct
            rs.MoveNext
        Loop
    Else
        MsgBox "No hay Datos para Mostrar", vbInformation, "Aviso!!!"
    End If
End Sub
Private Sub cmdAmortizar_Click()
    Set oIntangible = New dIntangible
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim rsCtas As ADODB.Recordset
    Set rsCtas = New ADODB.Recordset
    
    Dim lscod As String
    Dim lnMovNrox As Long
    Dim X As Integer
    
    Dim ldFechaAmortiza As Date
    Dim DMov As DMov
    Set DMov = New DMov
'    Dim oCon As DConecta
'    Set oCon = New DConecta
    Dim lsMovNro As String
    Dim lnMovNro As Long
    Dim lsMovNroR As String
    Dim lnMovNroR As Long
    Dim lsCodAmort As String
    Dim lsCodOpeAmort As String
    Dim lsCtaD As String, lsCtaH As String
    Dim nCtaActivados As Integer
    Dim oPrevio As clsPrevioFinan
    Dim oImpr As NContImprimir
    Set oPrevio = New clsPrevioFinan
    Set oImpr = New NContImprimir
    
    Dim lsMoneda As String
    Dim lncont As Integer
    
    
    Dim Dope As DOperacion
    Set Dope = New DOperacion
    
    If feAmortizar.TextMatrix(1, 1) <> "" Then
        
        If MsgBox("¿ Está seguro de realizar las amortizaciones para las intangibles seleccionadas", vbYesNo, "Atención") = vbNo Then Exit Sub
        
        'codigo para validar fechas(amrotizaciones contables)
        If ((DateAdd("M", 1, ldFecIni) - 1) < gdFecSis) And chkEstadistico.value = False Then
            MsgBox "No se puede amortizar contablemente en la fecha.", vbInformation, "Aviso!!!"
            Exit Sub
        End If
        
        Dim I As Integer
        ReDim lsErrores(1 To 9, 0 To 0)
        
        nCtaActivados = 0
        X = 1
        
        For I = 1 To feAmortizar.Rows - 1
            If ((feAmortizar.TextMatrix(I, 7) = "N/A") And (feAmortizar.TextMatrix(I, 9) = ".") And (feAmortizar.TextMatrix(I, 13) = 0)) Then
                    ReDim Preserve lsErrores(1 To 9, X)
                    lscod = feAmortizar.TextMatrix(I, 1)
                    lnMovNrox = feAmortizar.TextMatrix(I, 10)
                    Set rs = oIntangible.ObtenerSiAmortPendiente(lnMovNrox, lscod)
                    
                    lsErrores(1, X) = feAmortizar.TextMatrix(I, 1)
                    lsErrores(2, X) = feAmortizar.TextMatrix(I, 2)
                    lsErrores(3, X) = feAmortizar.TextMatrix(I, 3)
                    lsErrores(4, X) = feAmortizar.TextMatrix(I, 4)
                    lsErrores(5, X) = feAmortizar.TextMatrix(I, 5)
                    lsErrores(6, X) = feAmortizar.TextMatrix(I, 6)
                    lsErrores(7, X) = feAmortizar.TextMatrix(I, 7)
                    lsErrores(8, X) = feAmortizar.TextMatrix(I, 8)
                    lsErrores(9, X) = IIf(rs!Estado = 0, "Amortización 100%", "Amortizaciones Pendientes")
                    
                    X = X + 1
            End If
            If ((feAmortizar.TextMatrix(I, 7) <> "N/A") And (feAmortizar.TextMatrix(I, 9) = ".") And (feAmortizar.TextMatrix(I, 13) = 0)) Then
                lscod = feAmortizar.TextMatrix(I, 1)
                lnMovNrox = feAmortizar.TextMatrix(I, 10)
                Set rs = oIntangible.ObtenerSiAmortPendiente(lnMovNrox, lscod)
                If rs!Estado = 0 Then
                     ReDim Preserve lsErrores(1 To 9, X)
                    lsErrores(1, X) = feAmortizar.TextMatrix(I, 1)
                    lsErrores(2, X) = feAmortizar.TextMatrix(I, 2)
                    lsErrores(3, X) = feAmortizar.TextMatrix(I, 3)
                    lsErrores(4, X) = feAmortizar.TextMatrix(I, 4)
                    lsErrores(5, X) = feAmortizar.TextMatrix(I, 5)
                    lsErrores(6, X) = feAmortizar.TextMatrix(I, 6)
                    lsErrores(7, X) = feAmortizar.TextMatrix(I, 7)
                    lsErrores(8, X) = feAmortizar.TextMatrix(I, 8)
                    lsErrores(9, X) = IIf(rs!Estado = 0, "Amortización 100%.", "Amortizaciones Pendientes.")
                    X = X + 1
                End If
'                If Not rs.EOF And Not rs.BOF Then
'
'                    ReDim Preserve lsErrores(1 To 9, x)
'                    lsErrores(1, x) = feAmortizar.TextMatrix(i, 1)
'                    lsErrores(2, x) = feAmortizar.TextMatrix(i, 2)
'                    lsErrores(3, x) = feAmortizar.TextMatrix(i, 3)
'                    lsErrores(4, x) = feAmortizar.TextMatrix(i, 4)
'                    lsErrores(5, x) = feAmortizar.TextMatrix(i, 5)
'                    lsErrores(6, x) = feAmortizar.TextMatrix(i, 6)
'                    lsErrores(7, x) = feAmortizar.TextMatrix(i, 7)
'                    lsErrores(8, x) = feAmortizar.TextMatrix(i, 8)
'                    lsErrores(9, x) = IIf(rs!Estado = 0, "Amortización 100%.", "Amortizaciones Pendientes.")
'                    x = x + 1
'
'                End If
                nCtaActivados = nCtaActivados + 1
            End If
            
        Next I
        
        If UBound(lsErrores, 2) > 0 Then
            MsgBox "No se ha realizado ninguna amortización por que existen errores", vbInformation, "Aviso!!!"
            MostrarErrores
            Exit Sub
        Else
            If nCtaActivados > 0 Then
                DMov.BeginTrans
            'Codigo para el Asiento Contable
                   
                    'ldFechaAmortiza = DateAdd("M", 1, CDate("01/" & Format(cboMes.ListIndex + 1, "00") & "/" & txtAnio.Text)) - 1
                    'ldFechaAmortiza = DateAdd("M", 1, CDate("01/" & Format(DatePart("M", ldFecIni), "00") & "/" & Format(DatePart("Y", ldFecIni), "00"))) - 1
                    'ldFechaAmortiza = DateAdd("M", 1, ldFecIni) - 1
                    ldFechaAmortiza = DateAdd("M", 1, ldFecIni) - 1
                    'lsMovNro = DMov.GeneraMovNro(ldFechaAmortiza, gsCodAge, gsCodUser)
                    lsCodAmort = feAmortizar.TextMatrix(1, 1)
                
                    Select Case Mid(lsCodAmort, 7, 2)
                        Case "01"
                            lsCodOpeAmort = gAmortizaIntangibleLicencia
                        Case "02"
                            lsCodOpeAmort = gAmortizaIntangibleSoftware
                        Case "03"
                            lsCodOpeAmort = gAmortizaIntangibleOtros
                    End Select
                
                    Set rsCtas = Dope.ObtenerCtasAmortIntangible(lsCodOpeAmort)
                    lsCtaD = rsCtas!cCtaContCodD
                    lsCtaH = rsCtas!cCtaContCodH
                    X = 0
                    ReDim lnNroAsiento(1 To 3, 0 To 0)
                    For I = 1 To feAmortizar.Rows - 1
                        If ((feAmortizar.TextMatrix(I, 7) <> "N/A") And (feAmortizar.TextMatrix(I, 9) = ".") And (feAmortizar.TextMatrix(I, 13) = 0)) Then
                            X = X + 1
                            ReDim Preserve lnNroAsiento(1 To 3, X)
                            lsMovNro = DMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                            DMov.InsertaMov lsMovNro, lsCodOpeAmort, "Amortizacion Mensual de Intangibles ", 25
                            lnMovNro = DMov.GetnMovNro(lsMovNro)
                            'DMov.InsertaMovIntgAmort Left(lsMovNro, 4), feAmortizar.TextMatrix(i, 10), 1, feAmortizar.TextMatrix(i, 1), feAmortizar.TextMatrix(i, 7), lnMovNro, IIf(chkEstadistico.value = 1, 1, 0), gdFecSis
                            DMov.InsertaMovIntgAmort Left(lsMovNro, 4), feAmortizar.TextMatrix(I, 10), 1, feAmortizar.TextMatrix(I, 1), feAmortizar.TextMatrix(I, 7), lnMovNro, IIf(chkEstadistico.value = 1, 1, 0), ldFechaAmortiza
                            DMov.InsertaMovCta lnMovNro, 1, Replace(lsCtaD, "AG", Right(feAmortizar.TextMatrix(I, 12), 2)), Round(feAmortizar.TextMatrix(I, 8), 2) 'Reemplazar AG falta
                            lnNroAsiento(1, X) = lnMovNro 'numero de movimiento
                            lnNroAsiento(2, X) = Round(feAmortizar.TextMatrix(I, 8), 2) 'monto a amortizar
                            lnNroAsiento(3, X) = Right(feAmortizar.TextMatrix(I, 12), 2) 'agencia de la cuenta
                            
                            DMov.InsertaMovCta lnNroAsiento(1, X), 2, Replace(lsCtaH, "AG", lnNroAsiento(3, X)), lnNroAsiento(2, X) * -1
                        End If
                    Next I
                    lncont = I - 1
'                    For i = 1 To UBound(lnNroAsiento, 2)
'                        DMov.InsertaMovCta lnNroAsiento(1, i), lncont, Replace(lsCtaH, "AG", lnNroAsiento(3, i)), lnNroAsiento(2, i) * -1
'                        lncont = lncont + 1
'                    Next i
'
'                    If chkEstadistico.value = 0 Then
'
'                            lsMovNroR = DMov.GeneraMovNro(ldFechaAmortiza, gsCodAge, gsCodUser)
'                            DMov.InsertaMov lsMovNroR, lsCodOpeAmort, "Amortizacion Mensual de Intangibles - Resumen " & lsRubroc, 10
'
'                            lnMovNroR = DMov.GetnMovNro(lsMovNroR)
'                            DMov.GeneraAsientoRes lnMovNroR, lnMovNro
'                            DMov.InsertaMovRef lnMovNroR, lnMovNro
'                            lsCadenaAsiento = oImpr.ImprimeAsientoContable(lsMovNroR, 66, 79)
'
'                    End If
                    ReDim Preserve lsNroAsientoRef(1 To 1, 0 To 0)
                    If chkEstadistico.value = 0 Then
                        For I = 1 To UBound(lnNroAsiento, 2)
                            ReDim Preserve lsNroAsientoRef(1 To 1, I)
                            lsMovNroR = DMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                            'DMov.InsertaMov lsMovNroR, lsCodOpeAmort, "Amortizacion " & lsRubroc & Format(DatePart("M", ldFecIni), "00") & Format(DatePart("YYYY", ldFecIni), "0000"), 10
                            DMov.InsertaMov lsMovNroR, lsCodOpeAmort, "Amortizacion " & lsRubroc & Fecha(DatePart("M", ldFecIni)) & " " & Format(DatePart("YYYY", ldFecIni), "0000"), 10
                            
                            lnMovNroR = DMov.GetnMovNro(lsMovNroR)
                            DMov.GeneraAsientoRes lnMovNroR, CLng(lnNroAsiento(1, I))
                            DMov.InsertaMovRef lnMovNroR, lnNroAsiento(1, I)
                            lsNroAsientoRef(1, I) = lsMovNroR
                        Next I
                    End If
            
            DMov.CommitTrans
                Me.Caption = "CONTABILIDAD: AMORTIZACIÓN DE INTANGIBLES"
                MsgBox "Amortización " + IIf(chkEstadistico.value = 0, "Contable con Asiento ", "Contable solo estadístico ") + "se generó OK.", vbInformation + vbOKOnly, "Atención"
            
'                If chkEstadistico.value = 0 Then
'                    'oPrevio.Show oImpr.ImprimeAsientoContable(lsMovNroR, 66, 79), lsCodOpeAmort, False, 66, gImpresora
'                    oPrevio.Show lsCadenaAsiento, lsCodOpeAmort, False, 66, gImpresora
'                End If
                lsCadenaAsiento = ""
                If chkEstadistico.value = 0 Then
                    For I = 1 To UBound(lsNroAsientoRef, 2)
                        lsCadenaAsiento = lsCadenaAsiento + oImpr.ImprimeAsientoContable(lsNroAsientoRef(1, I), 66, 79)
                    Next I
                    oPrevio.Show lsCadenaAsiento, lsCodOpeAmort, False, 66, gImpresora
                End If
                cboMes.ListIndex = Month(gdFecSis) - 1
                txtAnio.Text = Year(gdFecSis)
                CargarDatos
            Else
                MsgBox "No se ha realizado ninguna amortización, asegurese de haber marcado al menos una intangible", vbInformation, "Aviso!!"
            End If
        End If
    Else
        MsgBox "No hay Intangibles para amortizar", vbInformation, "Aviso!!!"
    End If
End Sub
Private Function Fecha(ByVal nMes As Integer) As String
    Select Case nMes
        Case 1
                Fecha = "Enero"
        Case 2
                Fecha = "Febrero"
        Case 3
                Fecha = "Marzo"
        Case 4
                Fecha = "Abril"
        Case 5
                Fecha = "Mayo"
        Case 6
                Fecha = "Junio"
        Case 7
                Fecha = "Julio"
        Case 8
                Fecha = "Agosto"
        Case 9
                Fecha = "Setiembre"
        Case 10
                Fecha = "Octubre"
        Case 11
                Fecha = "Noviembre"
        Case 12
                Fecha = "Diciembre"
    End Select
End Function
Private Sub MostrarErrores()
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExcelOpen As Boolean
    Dim lsArchivo As String
    Dim lnfil As Integer
    Dim lnCol As Integer
    Dim I As Integer, X As Integer
    lnfil = 4
    lnCol = 2
    
    Dim oMov As DMov
    Set oMov = New DMov
    
    'lsArchivo = App.path & "\Spooler\ErrAmortIntang_" & txtAnio.Text & Format(cboMes.ListIndex + 1, "00") & ".xls"
    lsArchivo = App.path & "\Spooler\ErrAmortIntang_" & oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser) & ".xls"
    If Not ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False) Then
        Exit Sub
    End If
    lbExcelOpen = True
        
    ExcelAddHoja "Error", xlLibro, xlHoja1, False
    xlAplicacion.Range("A1:R100").Font.Size = 10
    xlHoja1.Cells(2, 2) = "Error de Amortización de Intangibles"
        
    'Cabecera
    xlHoja1.Cells(3, 2) = "Codigo"
    xlHoja1.Cells(3, 3) = "Descripción"
    xlHoja1.Cells(3, 4) = "Rubro"
    xlHoja1.Cells(3, 5) = "Moneda"
    xlHoja1.Cells(3, 6) = "Valor"
    xlHoja1.Cells(3, 7) = "Valor MN"
    xlHoja1.Cells(3, 8) = "Nº de Amortizacion"
    xlHoja1.Cells(3, 9) = "Monto a Amortizar"
    xlHoja1.Cells(3, 10) = "Motivo de Rechazo"
    
    'Contenido
    For I = 1 To UBound(lsErrores, 2)
        For X = 1 To UBound(lsErrores, 1)
            If X = 1 Then
                xlHoja1.Range(xlHoja1.Cells(lnfil, lnCol), xlHoja1.Cells(lnfil, lnCol)).NumberFormat = "@"
            End If
            xlHoja1.Cells(lnfil, lnCol) = lsErrores(X, I)
            lnCol = lnCol + 1
        Next X
        lnCol = 2
        lnfil = lnfil + 1
    Next I
    
    ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
    CargaArchivo lsArchivo, App.path & "\SPOOLER"
    lbExcelOpen = False
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    cboMes.ListIndex = Month(gdFecSis) - 1
    txtAnio.Text = Year(gdFecSis)
    CargaCombo
End Sub
Private Sub CargaCombo()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oIntangible = New dIntangible
    Set rs = oIntangible.ListaTipoIntangible()
    If Not rs.EOF Then
        cboRubro.Clear
        Do While Not rs.EOF
            cboRubro.AddItem Trim(rs(1) & Space(100) & Trim(rs(0)))
            rs.MoveNext
        Loop
    End If
    If cboRubro.ListCount > 0 Then
        cboRubro.ListIndex = 1
    End If
End Sub

