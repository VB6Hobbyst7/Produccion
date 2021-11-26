VERSION 5.00
Begin VB.Form frmIntangibleExtorno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Intangibles - Extorno de Amortización"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16290
   Icon            =   "frmIntangibleExtorno.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   16290
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtGlosaExtorno 
      Height          =   315
      Left            =   1080
      TabIndex        =   10
      Top             =   4320
      Width           =   12255
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   14760
      TabIndex        =   8
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Frame frmBuscar 
      Caption         =   "Buscar"
      Height          =   975
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin VB.ComboBox cboRubro 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   2775
      End
      Begin VB.ComboBox cboMes 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmIntangibleExtorno.frx":030A
         Left            =   3600
         List            =   "frmIntangibleExtorno.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtAnio 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5640
         MaxLength       =   4
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdBuscar 
         Appearance      =   0  'Flat
         Caption         =   "Mostrar"
         Height          =   345
         Left            =   6720
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Rubro:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
   End
   Begin Sicmact.FlexEdit feExtorno 
      Height          =   2775
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   15975
      _ExtentX        =   28178
      _ExtentY        =   4895
      Cols0           =   17
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   $"frmIntangibleExtorno.frx":039A
      EncabezadosAnchos=   "300-1500-6000-1500-1000-1200-1200-1200-1200-350-0-0-0-0-0-0-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-9-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-4-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "L-L-L-C-C-R-C-R-R-C-L-C-C-L-R-L-C"
      FormatosEdit    =   "0-1-1-1-1-2-2-3-2-0-1-1-1-1-3-1-1"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "Extornar"
      Height          =   375
      Left            =   13560
      TabIndex        =   7
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Glosa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   4320
      Width           =   735
   End
End
Attribute VB_Name = "frmIntangibleExtorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**-------------------------------------------------------------------------------------**'
'** Formulario : frmIntangibleExtorno                                                   **'
'** Finalidad  : Este formulario permite realizar el extorno de las                     **'
'**             amortizaciones realizadas a las intangibles activadas                   **'
'** Programador: Paolo Hector Sinti Cabrera - PASI                                      **'
'** Fecha/Hora : 20140305 11:50 AM                                                      **'
'**-------------------------------------------------------------------------------------**'
Option Explicit
Dim oIntangible As dIntangible
Dim ldFecIni As Date
Dim ldFecFin As Date
Dim lsRubroc As String
Dim lsErrores() As String
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
    
    FormateaFlex feExtorno
    Set rs = oIntangible.ListaAmortizacionesparaExtorno(lsRubro, ldFecIni, ldFecFin)
    If Not rs.EOF Then
        Do While Not rs.EOF
            feExtorno.AdicionaFila
            row = feExtorno.row
            feExtorno.TextMatrix(row, 1) = rs!Codigo
            feExtorno.TextMatrix(row, 2) = rs!Descripcion
            feExtorno.TextMatrix(row, 3) = rs!Rubro
            feExtorno.TextMatrix(row, 4) = rs!Moneda
            feExtorno.TextMatrix(row, 5) = Format(rs!Valor, "#,#0.00")
            feExtorno.TextMatrix(row, 6) = Format(rs!ValorMN, "#,#0.00")
            feExtorno.TextMatrix(row, 7) = rs!NMesAmort
            feExtorno.TextMatrix(row, 8) = Format(rs!MontoAmor, "#,#0.00")
            feExtorno.TextMatrix(row, 9) = rs!Estado
            feExtorno.TextMatrix(row, 10) = rs!nMovNro
            feExtorno.TextMatrix(row, 11) = rs!cMovNro
            feExtorno.TextMatrix(row, 12) = rs!CtaCont
            feExtorno.TextMatrix(row, 13) = rs!FechaAct
            feExtorno.TextMatrix(row, 14) = rs!nEstaCont '0 = Estadistic ; 1 = COntable
            feExtorno.TextMatrix(row, 15) = rs!FechaAmort 'Fecha de la amortizacion realizada
            feExtorno.TextMatrix(row, 16) = rs!nmovValida
            rs.MoveNext
        Loop
    Else
        MsgBox "No hay Datos para Mostrar", vbInformation, "Aviso!!!"
    End If
End Sub
Private Sub cmdBuscar_Click()
    CargarDatos
End Sub
Private Sub cmdExtornar_Click()
    Dim rsamp As ADODB.Recordset
    Set rsamp = New ADODB.Recordset
    Dim rsCtas  As ADODB.Recordset
    Set rsCtas = New ADODB.Recordset
    
    Set oIntangible = New dIntangible
    Dim x As Integer
    Dim nCtaActivados As Integer
    Dim lsCodOpeAmort As String
    
    Dim nEstCont As Integer
    Dim nMovNro As Long
    Dim sIntgCod As String
    Dim sFeAmort As String
    'Dim ldFechaExtorno As Date
    Dim lsCodAmort As String
    
    
    Dim DMov As DMov
    Set DMov = New DMov
    Dim Dope As Doperacion
    Set Dope = New Doperacion
    
    Dim lsMovNro As String
    
    If feExtorno.TextMatrix(1, 1) <> "" Then
        If MsgBox("¿ Está seguro de realizar el extorno para las amortizaciones seleccionadas", vbYesNo, "Atención") = vbNo Then Exit Sub
        If Len(Me.txtGlosaExtorno.Text) = 0 Then
            MsgBox "No se ha ingresado una descripción válida para el extorno.", vbInformation, "Aviso!!!"
            Exit Sub
        End If
        
        Dim I As Integer
        ReDim lsErrores(1 To 9, 0 To 0)
        
        nCtaActivados = 0
        x = 1
        For I = 1 To feExtorno.Rows - 1
            If feExtorno.TextMatrix(I, 9) = "." Then 'Obtiene todos los checkeados
                nEstCont = feExtorno.TextMatrix(I, 14)
                'nMovNro = feExtorno.TextMatrix(i, 10)
                nMovNro = feExtorno.TextMatrix(I, 16)
                sIntgCod = feExtorno.TextMatrix(I, 1)
                sFeAmort = feExtorno.TextMatrix(I, 15)
                Set rsamp = oIntangible.ObtenerAmortizacionesPosteriores(nMovNro, sIntgCod, DateAdd("M", 1, ldFecIni), DateAdd("D", -1, DateAdd("M", 2, ldFecIni)))
                If Not rsamp.EOF And Not rsamp.BOF Then
                    ReDim Preserve lsErrores(1 To 9, x)
                    lsErrores(1, x) = feExtorno.TextMatrix(I, 1)
                    lsErrores(2, x) = feExtorno.TextMatrix(I, 2)
                    lsErrores(3, x) = feExtorno.TextMatrix(I, 3)
                    lsErrores(4, x) = feExtorno.TextMatrix(I, 4)
                    lsErrores(5, x) = feExtorno.TextMatrix(I, 5)
                    lsErrores(6, x) = feExtorno.TextMatrix(I, 6)
                    lsErrores(7, x) = feExtorno.TextMatrix(I, 7)
                    lsErrores(8, x) = feExtorno.TextMatrix(I, 8)
                    lsErrores(9, x) = "Cuenta con Amortizaciones Posteriores."
                    x = x + 1
                End If
                If rsamp.EOF And rsamp.BOF Then
                    If nEstCont = 1 Then
                        If DateDiff("D", CDate(sFeAmort), gdFecSis) > 0 Then
                            ReDim Preserve lsErrores(1 To 9, x)
                            lsErrores(1, x) = feExtorno.TextMatrix(I, 1)
                            lsErrores(2, x) = feExtorno.TextMatrix(I, 2)
                            lsErrores(3, x) = feExtorno.TextMatrix(I, 3)
                            lsErrores(4, x) = feExtorno.TextMatrix(I, 4)
                            lsErrores(5, x) = feExtorno.TextMatrix(I, 5)
                            lsErrores(6, x) = feExtorno.TextMatrix(I, 6)
                            lsErrores(7, x) = feExtorno.TextMatrix(I, 7)
                            lsErrores(8, x) = feExtorno.TextMatrix(I, 8)
                            lsErrores(9, x) = "Amortización Contable en dias anteriores."
                            x = x + 1
                        End If
                    End If
                End If
                nCtaActivados = nCtaActivados + 1
            End If
        Next I
        
        If UBound(lsErrores, 2) > 0 Then
            MsgBox "No se ha realizado ninguna extorno por que existen errores", vbInformation, "Aviso!!!"
            MostrarErrores
            Exit Sub
        Else
            If nCtaActivados > 0 Then
                
                DMov.BeginTrans
                    
                    lsCodAmort = feExtorno.TextMatrix(1, 1)
                    Select Case Mid(lsCodAmort, 7, 2)
                        Case "01"
                            lsCodOpeAmort = gAmortizaIntangibleLicencia
                        Case "02"
                            lsCodOpeAmort = gAmortizaIntangibleSoftware
                        Case "03"
                            lsCodOpeAmort = gAmortizaIntangibleOtros
                    End Select
'                    Set rsCtas = DOpe.ObtenerCtasAmortIntangible(lsCodOpeAmort)
'                    lsCtaD = rsCtas!cCtaContCodD
'                    lsCtaH = rsCtas!cCtaContCodH
                    
                    x = 0
                    
                    For I = 1 To feExtorno.Rows - 1
                       If feExtorno.TextMatrix(I, 9) = "." Then
                            nEstCont = feExtorno.TextMatrix(I, 14)
                            nMovNro = feExtorno.TextMatrix(I, 10)
                            sIntgCod = feExtorno.TextMatrix(I, 1)
                            sFeAmort = feExtorno.TextMatrix(I, 15)
                            lsMovNro = DMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                            If nEstCont = 1 Then 'contable
                                'oIntangible.ExtornaAmortizacion lsMovNro, gExtornoAmortizacion, Trim(Me.txtGlosaExtorno.Text), nMovNro
                                DMov.ExtornaMovimiento lsMovNro, nMovNro, gExtornoAmortizacion, Trim(Me.txtGlosaExtorno.Text)
                            Else 'Estadistico
                                DMov.ExtornaMovimiento lsMovNro, nMovNro, gExtornoAmortizacion, Trim(Me.txtGlosaExtorno.Text), True
                            End If
                       End If
                    Next I
                DMov.CommitTrans
                MsgBox "Se ha realizado con éxito los extornos. ", vbInformation, "Aviso!!!"
                CargarDatos
            Else
                MsgBox "No se ha realizado ningun extorno, asegurese de haber marcado al menos una amortización", vbInformation, "Aviso!!"
            End If
        End If
    Else
        MsgBox "No existen datos para extornar", vbInformation, "Aviso!!!"
    End If
End Sub
Private Sub MostrarErrores()
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExcelOpen As Boolean
    Dim lsArchivo As String
    Dim lnfil As Integer
    Dim lnCol As Integer
    Dim I As Integer, x As Integer
    lnfil = 4
    lnCol = 2
    
    Dim oMov As DMov
    Set oMov = New DMov
    
    'lsArchivo = App.path & "\Spooler\ErrAmortIntang_" & txtAnio.Text & Format(cboMes.ListIndex + 1, "00") & ".xls"
    lsArchivo = App.path & "\Spooler\ErrExtornoIntang_" & oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser) & ".xls"
    If Not ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False) Then
        Exit Sub
    End If
    lbExcelOpen = True
        
    ExcelAddHoja "Error", xlLibro, xlHoja1, False
    xlAplicacion.Range("A1:R100").Font.Size = 10
    xlHoja1.Cells(2, 2) = "Error de Extorno de Intangibles"
        
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
        For x = 1 To UBound(lsErrores, 1)
            If x = 1 Then
                xlHoja1.Range(xlHoja1.Cells(lnfil, lnCol), xlHoja1.Cells(lnfil, lnCol)).NumberFormat = "@"
            End If
            xlHoja1.Cells(lnfil, lnCol) = lsErrores(x, I)
            lnCol = lnCol + 1
        Next x
        lnCol = 2
        lnfil = lnfil + 1
    Next I
    
    ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
    CargaArchivo lsArchivo, App.path & "\SPOOLER"
    lbExcelOpen = False
End Sub
Private Sub Form_Load()
    cboMes.ListIndex = Month(gdFecSis) - 1
    txtAnio.Text = Year(gdFecSis)
    CargaCombo
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
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
        cboRubro.ListIndex = 0
    End If
End Sub
