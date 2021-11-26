VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLogMantSaldos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Saldos"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
   Icon            =   "frmLogMantSaldos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar P 
      Height          =   255
      Left            =   90
      TabIndex        =   17
      Top             =   8100
      Visible         =   0   'False
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.ComboBox cboTpoAlm 
      Height          =   315
      Left            =   4755
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   45
      Width           =   4380
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexE 
      Height          =   195
      Left            =   5565
      TabIndex        =   13
      Top             =   7830
      Visible         =   0   'False
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   344
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdExcell 
      Caption         =   "Imp Excel"
      Height          =   360
      Left            =   5790
      TabIndex        =   12
      Top             =   7665
      Width           =   1080
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   360
      Left            =   6945
      TabIndex        =   11
      Top             =   7665
      Width           =   1080
   End
   Begin VB.CommandButton cmdBuscarSig 
      Caption         =   "Buscar S&iguiente >>>"
      Height          =   360
      Left            =   3450
      TabIndex        =   10
      Top             =   7665
      Width           =   2085
   End
   Begin Sicmact.TxtBuscar txtAlmacen 
      Height          =   285
      Left            =   795
      TabIndex        =   8
      Top             =   405
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   503
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      sTitulo         =   ""
   End
   Begin MSComCtl2.DTPicker dtFecha 
      Height          =   330
      Left            =   795
      TabIndex        =   5
      Top             =   30
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   582
      _Version        =   393216
      Format          =   70713345
      CurrentDate     =   37287
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   360
      Left            =   2325
      TabIndex        =   4
      Top             =   7665
      Width           =   1080
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   8085
      TabIndex        =   3
      Top             =   7665
      Width           =   1080
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   360
      Left            =   1200
      TabIndex        =   2
      Top             =   7665
      Width           =   1080
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   360
      Left            =   75
      TabIndex        =   1
      Top             =   7665
      Width           =   1080
   End
   Begin Sicmact.FlexEdit Flex 
      Height          =   6795
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   11986
      Cols0           =   6
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "#-Ok-Codigo-Descripcion-Stock-Monto"
      EncabezadosAnchos=   "400-400-1200-4500-1000-1000"
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
      ColumnasAEditar =   "X-1-X-3-4-5"
      TextStyleFixed  =   3
      ListaControles  =   "0-4-0-0-0-0"
      EncabezadosAlineacion=   "R-C-L-L-R-R"
      FormatosEdit    =   "3-0-0-0-2-2"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      Appearance      =   0
      ColWidth0       =   405
      RowHeight0      =   300
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Almacén Tipo :"
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
      Height          =   180
      Index           =   0
      Left            =   3225
      TabIndex        =   16
      Top             =   105
      Width           =   1500
   End
   Begin VB.OLE OleExcel 
      Appearance      =   0  'Flat
      AutoActivate    =   3  'Automatic
      Enabled         =   0   'False
      Height          =   255
      Left            =   0
      SizeMode        =   1  'Stretch
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label lblAlmacenG 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2355
      TabIndex        =   9
      Top             =   405
      Width           =   6765
   End
   Begin VB.Label lblAmacen 
      Caption         =   "Almacen"
      Height          =   180
      Left            =   45
      TabIndex        =   7
      Top             =   465
      Width           =   720
   End
   Begin VB.Label lblFecha 
      Caption         =   "Fecha :"
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   75
      Width           =   630
   End
End
Attribute VB_Name = "frmLogMantSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim xlHoja2 As Excel.Worksheet


Dim lsCad As String
Dim lnI As Integer
Dim color As Integer 'variable para activar el color del seleccionado ANPS
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
'*******************************

Private Sub cmdBuscar_Click()
    
    lsCad = UCase(Trim(InputBox("Ingrese Codigo a Buscar.", "Aviso"))) 'anps
'
  If lsCad = "" Then Exit Sub
'COMENTADO ANPS
'    For lnI = 1 To Me.Flex.Rows - 1
'        'If InStr(1, UCase(Flex.TextMatrix(lnI, 2)), lsCad) <> 0 Then
'        If InStr(1, UCase(Flex.TextMatrix(lnI, 2)), lsCad) <> 0 Then 'JIPR20210221 CAMBIAR VALOR A 1
'            Flex.row = lnI
'            Exit Sub
'        End If
'    Next lnI
'Flex.rsFlex.Find
 ' ANPS FILTRO PARA BUSQUEDA
    color = 0
    GetLogStock lsCad, 1 'Carga la lista con filtro opcion 1 anps
        If Flex.TextMatrix(1, 2) = lsCad Then
            Flex.row = 1
            Flex.col = 1
            Flex.BackColorRow RGB(0, 120, 215) '("#0078d7") PONE EL COLOR A LA SELECCION
            color = 1
        End If
 'FIN ANPS
  Exit Sub
ErrMsg:
        MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdBuscarSig_Click() ' OPCION DESABILITADA

    If lsCad = "" Then Exit Sub
    
    For lnI = lnI + 1 To Me.Flex.Rows - 1
        'If InStr(1, UCase(Flex.TextMatrix(lnI, 2)), lsCad) <> 0 Then
        If InStr(1, UCase(Flex.TextMatrix(lnI, 2)), lsCad) <> 0 Then 'JIPR20210221 CAMBIAR VALOR A 1  ANPS
            Flex.row = lnI
            Exit Sub
        End If
        
    Next lnI
End Sub

Private Sub cmdEditar_Click()
    Me.cmdEditar.Enabled = False
    Me.Flex.lbEditarFlex = True
End Sub

Private Sub cmdExcell_Click()
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    
    Dim lbLibroOpen As Boolean
    Dim lsArchivoN As String
    
    Dim lnMoto As Currency
    Dim lnCantidad As Currency
    Dim lsCtaAnt As String
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    FlexE.Cols = 8
    FlexE.Rows = 1

    Me.FlexE.TextMatrix(0, 0) = "Codigo"
    Me.FlexE.TextMatrix(0, 1) = "Descripcion"
    Me.FlexE.TextMatrix(0, 2) = "Saldo"
    Me.FlexE.TextMatrix(0, 3) = "Cantidad"
    Me.FlexE.TextMatrix(0, 4) = "Cta.Cont"
    Me.FlexE.TextMatrix(0, 5) = "CtaCont.Descripcion"
    Me.FlexE.TextMatrix(0, 6) = "Tot.Sal."
    Me.FlexE.TextMatrix(0, 7) = "Tot.Cant."

    Set rs = oALmacen.CargaStock(CDate(Me.dtFecha), Val(Right(Me.cboTpoAlm.Text, 5)), Me.txtAlmacen.Text)

    Me.Height = Me.Height + 300
    Me.P.Visible = True
    Me.P.Min = 0
    Me.P.value = 0
    Me.P.Max = rs.RecordCount + 10

    lnMoto = 0
    lnCantidad = 0



    While Not rs.EOF
        FlexE.Rows = FlexE.Rows + 1
        If lsCtaAnt <> rs!Cta And FlexE.Rows <> 2 Then
            Me.FlexE.TextMatrix(FlexE.Rows - 2, 6) = Format(lnMoto, "#0.00")
            Me.FlexE.TextMatrix(FlexE.Rows - 2, 7) = Format(lnCantidad, "#0.00")
            lnMoto = 0
            lnCantidad = 0
        End If

        Me.FlexE.TextMatrix(FlexE.Rows - 1, 0) = rs!cBSCod
        Me.FlexE.TextMatrix(FlexE.Rows - 1, 1) = rs!cBSDescripcion
        Me.FlexE.TextMatrix(FlexE.Rows - 1, 2) = rs!nMonto
        Me.FlexE.TextMatrix(FlexE.Rows - 1, 3) = rs!nStock
        Me.FlexE.TextMatrix(FlexE.Rows - 1, 4) = rs!Cta
        Me.FlexE.TextMatrix(FlexE.Rows - 1, 5) = rs!cCtaContDesc
        lnMoto = lnMoto + rs!nMonto
        lnCantidad = lnCantidad + rs!nStock
        lsCtaAnt = rs!Cta
        DoEvents
        Me.P.value = Me.P.value + 1

        rs.MoveNext
    Wend

    Me.FlexE.TextMatrix(FlexE.Rows - 1, 6) = Format(lnMoto, "#0.00")
    Me.FlexE.TextMatrix(FlexE.Rows - 1, 7) = Format(lnCantidad, "#0.00")
    lnMoto = 0
    lnCantidad = 0


    lsArchivoN = App.path & "\Spooler\" & Format(CDate(Me.dtFecha), "yyyymmdd") & ".xls"

    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis, "yyyymmdd") & "_" & gsCodUser, xlLibro, xlHoja1
       Call GeneraReporte
       OleExcel.Class = "ExcelWorkSheet"
       ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1

       OleExcel.SourceDoc = lsArchivoN
       OleExcel.Verb = 1
       OleExcel.Action = 1
       OleExcel.DoVerb -1
    End If

        'ARLO 20160126 ***
        gsOpeCod = LogPistaMantemientoSaldos
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Imprio en Excel el Saldo a la fecha " & CDate(Me.dtFecha)
        Set objPista = Nothing
        '**************

    Me.Height = Me.Height - 300
    Me.P.Visible = False
    Me.P.Min = 0
    Me.P.value = 0
End Sub

Private Sub cmdGrabar_Click()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    
    If MsgBox("Desea Grabar las operaciones del día indicado ? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Set rs = Me.Flex.GetRsNew
    
    oALmacen.InsertaBSSaldosDia Me.txtAlmacen.Text, rs, CDate(Me.dtFecha), Val(Right(Me.cboTpoAlm.Text, 5))
    
    GetLogStock
    
    MsgBox "Grabación completa.", vbInformation, "Aviso"
    'ARLO 20160126 ***
    Dim lsPalabras As String
    If (Me.txtAlmacen.Text = 1) Then
    lsPalabras = "Almacen Principal"
    ElseIf (Me.txtAlmacen.Text = 2) Then
    lsPalabras = "Almacen Huanuco"
    ElseIf (Me.txtAlmacen.Text = 3) Then
    lsPalabras = "Almacen Pucallpa"
    ElseIf (Me.txtAlmacen.Text = 4) Then
    lsPalabras = "Almacen Calle Arequipa"
    ElseIf (Me.txtAlmacen.Text = 6) Then
    lsPalabras = "Almacen Yurimaguas"
    ElseIf (Me.txtAlmacen.Text = 7) Then
    lsPalabras = "Almacen Tingo Maria"
    ElseIf (Me.txtAlmacen.Text = 9) Then
    lsPalabras = "Almacen Belen"
    ElseIf (Me.txtAlmacen.Text = 10) Then
    lsPalabras = "Almacen Tarapoto"
    ElseIf (Me.txtAlmacen.Text = 12) Then
    lsPalabras = "Almacen Aguaytia"
    ElseIf (Me.txtAlmacen.Text = 13) Then
    lsPalabras = "Almacen Requena"
    ElseIf (Me.txtAlmacen.Text = 24) Then
    lsPalabras = "Almacen Cajamarca"
    ElseIf (Me.txtAlmacen.Text = 31) Then
    lsPalabras = "Almacen Punchana"
    ElseIf (Me.txtAlmacen.Text = 33) Then
    lsPalabras = "Almacen Minka"
    ElseIf (Me.txtAlmacen.Text = 37) Then
    lsPalabras = "Almacen San Juan"
    ElseIf (Me.txtAlmacen.Text = 38) Then
    lsPalabras = "Almacen Moyobamba"
    ElseIf (Me.txtAlmacen.Text = 39) Then
    lsPalabras = "Almacen Tocache"
    End If
    gsOpeCod = LogPistaMantemientoSaldos
    Set objPista = New COMManejador.Pista
    objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Grabo Mantenimiento de Saldo a la fecha " & CDate(Me.dtFecha) & " en el " & lsPalabras
    Set objPista = Nothing
    '**************
End Sub

Private Sub cmdImprimir_Click()
    Dim oPrevio As clsPrevio
    Set oPrevio = New clsPrevio
    
    Dim lsCorr As String * 8
    Dim lsCodigo As String * 20
    Dim lsDescripcion As String * 50
    Dim lsStock As String * 10
    Dim lsMonto As String * 15
    
    Dim lsCadena As String
    Dim lsCadenaCab As String
    Dim lnI As Long
    Dim lnIndice As Long
    
    lnI = 0
    
    lsCadenaCab = lsCadenaCab & oImpresora.gPrnCondensadaON
    
    RSet lsCorr = Format("Nro", "00000")
    lsCodigo = Flex.TextMatrix(lnI, 2) 'ANPS
    lsDescripcion = Flex.TextMatrix(lnI, 3) 'ANPS
    lsStock = Format(Flex.TextMatrix(lnI, 4), "#,##0.00") 'ANPS
    lsMonto = Format(Flex.TextMatrix(lnI, 5), "#,##0.00") 'ANPS

    lsCadenaCab = Me.dtFecha & oImpresora.gPrnSaltoLinea
    lsCadenaCab = lsCadenaCab & lsCorr & Space(10) & lsCodigo & Space(2) & lsDescripcion & Space(2) & lsStock & Space(2) & lsMonto & oImpresora.gPrnSaltoLinea & String(Len(lsCorr & Space(2) & lsCodigo & Space(2) & lsDescripcion & Space(2) & lsStock & Space(2) & lsMonto), "=") & oImpresora.gPrnSaltoLinea
    
  lsCadena = lsCadenaCab
    
    For lnI = 1 To Me.Flex.Rows - 1
        RSet lsCorr = Format(Flex.TextMatrix(lnI, 0), "00000")
        RSet lsCodigo = Flex.TextMatrix(lnI, 2)  'ANPS
        lsDescripcion = Flex.TextMatrix(lnI, 3) 'ANPS
        RSet lsStock = Format(Flex.TextMatrix(lnI, 4), "#,##0.00") 'ANPS
        RSet lsMonto = Format(Flex.TextMatrix(lnI, 5), "#,##0.00") 'ANPS
        
        lnIndice = lnIndice + 1
        lsCadena = lsCadena & lsCorr & Space(2) & lsCodigo & Space(2) & lsDescripcion & Space(2) & lsStock & Space(2) & lsMonto & oImpresora.gPrnSaltoLinea
        
        If lnIndice = 54 Then
            lnIndice = 0
            lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
            lsCadena = lsCadena & lsCadenaCab & oImpresora.gPrnSaltoLinea
        End If
    Next lnI
    
    oPrevio.Show lsCadena, "Listado de Bienes (STOCK)", True, 66
    
        'ARLO 20160126 ***
        gsOpeCod = LogPistaMantemientoSaldos
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, gIngresarSistema, "Se Imprimio Mantenimiento de Saldo a la fecha " & CDate(Me.dtFecha)
        Set objPista = Nothing
        '**************
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub dtFecha_Change()
    GetLogStock
End Sub
Private Sub Flex_Click() ' ANPS PARA DESACTIVAR EL COLOR DEL SELECCIONADO
    
    If color = 1 Then
        Flex.row = 1
        Flex.col = 1
        Flex.BackColorRow (vbWhite)
        color = 0
    End If
    
End Sub

Private Sub Form_Load()
    Dim oDoc As DOperaciones
    Set oDoc = New DOperaciones
    Dim oCon As DConstantes
    Set oCon = New DConstantes
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Me.cmdBuscarSig.Visible = False 'anps
    
    Me.dtFecha = Format(gdFecSis, gsFormatoFechaView)
    
    Set rs = oCon.GetConstante(5010, False)
    CargaCombo rs, cboTpoAlm
    cboTpoAlm.ListIndex = 0
    
    Me.txtAlmacen.rs = oDoc.GetAlmacenes
    
    Me.txtAlmacen.Text = "1"
    Me.lblAlmacenG.Caption = txtAlmacen.psDescripcion
    
    GetLogStock
End Sub

Private Sub txtAlmacen_EmiteDatos()
    lblAlmacenG.Caption = txtAlmacen.psDescripcion
End Sub

Private Sub GetLogStock(Optional ByVal lcal As String = "", Optional ByVal cod As Integer = 0) 'ANPS
    Dim oAlmacen As DLogAlmacen
    Set oAlmacen = New DLogAlmacen
  
    If Me.txtAlmacen.Text = "" Then
        MsgBox "Debe ingresar un almacen.", vbInformation, "Aviso"
        txtAlmacen.SetFocus
        Exit Sub
    End If

    Flex.rsFlex = oAlmacen.GetLogAlmacenStock(Me.txtAlmacen.Text, Me.dtFecha, Val(Right(Me.cboTpoAlm.Text, 5)), lsCad, cod)

    Me.cmdEditar.Enabled = True
    Me.Flex.lbEditarFlex = False

End Sub

Private Sub GeneraReporte()
    Dim I As Integer
    Dim k As Integer
    Dim J As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim sConec As String
    Dim lnAcum As Currency
    Dim VSQL As String
    
    Dim lnFilaMarcaIni As Integer
    Dim lnFilaMarcaFin As Integer
    
    Dim sTipoGara As String
    Dim sTipoCred As String
   
    lnFilaMarcaIni = 1
    
    xlHoja1.Columns.Range("A:A").Select
    xlHoja1.Columns.Range("A:A").NumberFormat = "@"
 
    P.value = 0
    For I = 0 To Me.FlexE.Rows - 1
        lnAcum = 0
        For J = 0 To Me.FlexE.Cols - 1
            xlHoja1.Cells(I + 1, J + 1) = Me.FlexE.TextMatrix(I, J)
            If I > 1 And J > 1 Then
                
                If IsNumeric(Me.FlexE.TextMatrix(I, J)) Then
                    lnAcum = lnAcum + CCur(Me.FlexE.TextMatrix(I, J))
                End If
            End If
        Next J
        
    If Me.FlexE.TextMatrix(i, 0) <> "" And i > 0 Then
            xlHoja1.Range("A" & lnFilaMarcaIni & ":H" & Trim(Str(i + 0))).Select
        
            xlHoja1.Range("A" & lnFilaMarcaIni & ":H" & Trim(Str(i + 0))).Borders(xlDiagonalDown).LineStyle = xlNone
            xlHoja1.Range("A" & lnFilaMarcaIni & ":H" & Trim(Str(i + 0))).Borders(xlDiagonalUp).LineStyle = xlNone
            With xlHoja1.Range("A" & lnFilaMarcaIni & ":H" & Trim(Str(i + 0))).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With xlHoja1.Range("A" & lnFilaMarcaIni & ":H" & Trim(Str(i + 0))).Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With xlHoja1.Range("A" & lnFilaMarcaIni & ":H" & Trim(Str(i + 0))).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With xlHoja1.Range("A" & lnFilaMarcaIni & ":H" & Trim(Str(i + 0))).Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            xlHoja1.Range("A" & lnFilaMarcaIni & ":H" & Trim(Str(i + 0))).Borders(xlInsideVertical).LineStyle = xlNone
            
            lnFilaMarcaIni = i + 1
        End If
        
        If i > 1 Then
            'xlHoja1.Range("A1:H" & Trim(Str(Me.flex.Rows))).Select
            'lnFilaMarcaIni = 0
            
            'VSQL = Format(lnAcum, "#,##0.00")  ' "=SUMA(" & Trim(ExcelColumnaString(3)) & Trim(I + 1) & ":" & Trim(ExcelColumnaString(Me.Flex.Cols)) & Trim(I + 1) & ")"
            'xlHoja1.Cells(I + 1, Me.Flex.Cols + 1).Formula = VSQL
            'xlHoja1.Cells(I + 1, Me.flex.Cols + 1) = VSQL
        End If
        DoEvents
        P.value = P.value + 1
    Next i
        
      xlHoja1.Range("A" & lnFilaMarcaIni & ":H" & Trim(Str(i + 0))).Select

    xlHoja1.Range("A" & lnFilaMarcaIni & ":H" & Trim(Str(i + 0))).Borders(xlDiagonalDown).LineStyle = xlNone
    xlHoja1.Range("A" & lnFilaMarcaIni & ":H" & Trim(Str(i + 0))).Borders(xlDiagonalUp).LineStyle = xlNone
    With xlHoja1.Range("A" & lnFilaMarcaIni & ":H" & Trim(Str(i + 0))).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A" & lnFilaMarcaIni & ":H" & Trim(Str(i + 0))).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A" & lnFilaMarcaIni & ":H" & Trim(Str(i + 0))).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A" & lnFilaMarcaIni & ":H" & Trim(Str(i + 0))).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    xlHoja1.Range("A" & lnFilaMarcaIni & ":H" & Trim(Str(i + 0))).Borders(xlInsideVertical).LineStyle = xlNone
    
    lnFilaMarcaIni = i + 1
	
    xlHoja1.Range("A1:A" & Trim(Str(Me.FlexE.Rows))).Font.Bold = True
    xlHoja1.Range("B1:B" & Trim(Str(Me.FlexE.Rows))).Font.Bold = True
    xlHoja1.Range("G1:G" & Trim(Str(Me.FlexE.Rows))).Font.Bold = True
    xlHoja1.Range("H1:H" & Trim(Str(Me.FlexE.Rows))).Font.Bold = True
    xlHoja1.Range("1:1").Font.Bold = True

    xlHoja1.Range("C2:D" & Trim(Str(Me.FlexE.Rows))).NumberFormat = "#,##0.00"
    xlHoja1.Range("G2:H" & Trim(Str(Me.FlexE.Rows))).NumberFormat = "#,##0.00"
    xlHoja1.Range("E2:E" & Trim(Str(Me.FlexE.Rows))).NumberFormat = "@"

    With xlHoja1.PageSetup
            .LeftHeader = ""
            .CenterHeader = "&""Arial,Negrita""&14LISTADO DE SALDOS " & Me.dtFecha
            .RightHeader = "&P"
            .LeftFooter = ""
            .CenterFooter = ""
            .RightFooter = ""
'            .LeftMargin = Application.InchesToPoints(0.21)
'            .RightMargin = Application.InchesToPoints(0)
'            .TopMargin = Application.InchesToPoints(0.41)
'            .BottomMargin = Application.InchesToPoints(0.27)
'            .HeaderMargin = Application.InchesToPoints(0.13)
'            .FooterMargin = Application.InchesToPoints(0)
            .PrintHeadings = False
            .PrintGridlines = False
            .PrintComments = xlPrintNoComments
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = xlPortrait
            .Draft = False
            '.PaperSize = xlPaperLetter
            .FirstPageNumber = xlAutomatic
            .Order = xlDownThenOver
            .BlackAndWhite = False
            .Zoom = 60
        End With
    xlHoja1.Cells.Select
    xlHoja1.Cells.EntireColumn.AutoFit
End Sub

Public Sub Inicio(pbParaSIG As Boolean)
    If pbParaSIG Then
        Me.cmdEditar.Visible = False
        Me.cmdGrabar.Visible = False
    End If
    Me.Show 1
End Sub

