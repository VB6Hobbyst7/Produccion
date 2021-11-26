VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLogAfTrans 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle de Bienes por Serie"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   Icon            =   "frmLogAfTrans.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMantSalidas 
      Caption         =   "&Mantenimien."
      Height          =   300
      Left            =   3240
      TabIndex        =   16
      Top             =   1755
      Width           =   1185
   End
   Begin VB.CommandButton cmdTrasferirActivoFijo 
      Caption         =   "&Transf AF"
      Height          =   300
      Left            =   4515
      TabIndex        =   15
      Top             =   1755
      Width           =   1185
   End
   Begin VB.CommandButton cmdReporteS 
      Caption         =   "&Reporte Sal"
      Height          =   300
      Left            =   3255
      TabIndex        =   14
      Top             =   1755
      Width           =   1185
   End
   Begin VB.CommandButton cmdProcesarS 
      Caption         =   "&Procesar Sal"
      Height          =   300
      Left            =   4515
      TabIndex        =   13
      Top             =   1755
      Width           =   1185
   End
   Begin VB.Frame fraFechas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   1635
      Left            =   30
      TabIndex        =   4
      Top             =   45
      Width           =   6915
      Begin Sicmact.TxtBuscar txtAlmacen 
         Height          =   315
         Left            =   210
         TabIndex        =   5
         Top             =   300
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
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
      Begin MSMask.MaskEdBox mskFecIni 
         Height          =   285
         Left            =   1110
         TabIndex        =   6
         Top             =   1170
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFecFin 
         Height          =   285
         Left            =   4290
         TabIndex        =   7
         Top             =   1185
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFecIni 
         Caption         =   "Fecha Inicial :  "
         Height          =   210
         Left            =   1140
         TabIndex        =   10
         Top             =   945
         Width           =   1110
      End
      Begin VB.Label lblFecFin 
         Caption         =   "Fecha Final"
         Height          =   210
         Left            =   4290
         TabIndex        =   9
         Top             =   945
         Width           =   1110
      End
      Begin VB.Label lblAmacenG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1620
         TabIndex        =   8
         Top             =   300
         Width           =   5160
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   315
      Left            =   5775
      TabIndex        =   3
      Top             =   1748
      Width           =   1170
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar Ing"
      Height          =   300
      Left            =   4515
      TabIndex        =   2
      Top             =   1755
      Width           =   1185
   End
   Begin VB.CommandButton cmdReporte 
      Caption         =   "&Reporte Ing"
      Height          =   300
      Left            =   3255
      TabIndex        =   0
      Top             =   1755
      Width           =   1185
   End
   Begin MSComctlLib.ProgressBar Prg 
      Height          =   180
      Left            =   60
      TabIndex        =   1
      Top             =   2115
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   318
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flex 
      Height          =   270
      Left            =   4560
      TabIndex        =   11
      Top             =   2130
      Visible         =   0   'False
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   476
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.OLE OleExcel 
      Appearance      =   0  'Flat
      AutoActivate    =   3  'Automatic
      Enabled         =   0   'False
      Height          =   255
      Left            =   0
      SizeMode        =   1  'Stretch
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   270
   End
End
Attribute VB_Name = "frmLogAfTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim xlHoja2 As Excel.Worksheet

Dim lsCaption As String
Dim lbIngreso As Boolean

Dim lbIngresoVista As Boolean
Dim lbSalidaVista As Boolean
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
Dim lsPalabras As String
'*******************************

Public Sub Ini(pbIngreso As Boolean, pbSalida As Boolean, psCaption As String)
    
    lbIngresoVista = pbIngreso
    lbSalidaVista = pbSalida
    lsCaption = psCaption
    
    Me.Show 1
End Sub

Private Sub cmdMantSalidas_Click()
    If Not IsDate(Me.mskFecIni.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        mskFecIni.SetFocus
        Exit Sub
    ElseIf Not IsDate(Me.mskFecFin.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        mskFecFin.SetFocus
        Exit Sub
    End If
    
    frmLogSalidasAFBND.Ini CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text)
End Sub

Private Sub cmdProcesar_Click()
    Dim lnItem As Long
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim lsCadena As String
    Dim lnPagina As Long
    
    Dim lsFecha As String
    Dim lsDocumento As String
    Dim lsDescripcion As String
    Dim lsBienes As String
    Dim lsCantidad As String
    Dim lsImporte As String
    Dim lsPersona As String
    
    Dim lsDocAnt As String
    Dim lnAcumulador As Currency
        
    Dim rsSerie As New ADODB.Recordset
    'Set rsSerie = New ADODB.Recordset
    
    Dim lnI As Integer
    
    Dim lsArchivoN  As String
    Dim lbLibroOpen As Boolean
    
    Dim lsSerie As String
    Dim lsProveedor As String
    Dim lsFactura As String
    
    Dim lnIGVAcum As Currency
    Dim lnValorAcum As Currency
    
    If Not IsDate(Me.mskFecIni.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        mskFecIni.SetFocus
        Exit Sub
    ElseIf Not IsDate(Me.mskFecFin.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        mskFecFin.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Desea procesar ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso") = vbNo Then Exit Sub
    
    Set rs = oALmacen.GetIngresosAlmacenAF(CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text), Me.txtAlmacen.Text)
    oALmacen.EliminaMovBsControl CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text), True
    
    Prg.Max = rs.RecordCount + 2
    Prg.value = Prg.Min
        
    While Not rs.EOF
        
        If rs!bSerie Then
            Set rsSerie = oALmacen.GetSerieMov(rs!cBSCod, rs!cDocNro, 42, rs!nMovItem)      'rs!nMovNro)
        End If
        
        lnIGVAcum = 0
        lnValorAcum = 0
        For lnI = 1 To rs!nMovCant
            'If rs!bSerie Or (rsSerie.EOF And rsSerie.BOF) Then
            If rs!bSerie Or (rsSerie.State = 0) Then
                lsSerie = Trim(Str(Year(gdFecSis))) & "-" & Format(CLng(oALmacen.GetCorrelaSerie(rs!cBSCod)) + 1, "00000000")
                oALmacen.InsertMovBsControl Year(gdFecSis), rs!nMovNro, rs!cBSCod, lsSerie, Format(rs!nMovImporte / rs!nMovCant, "#.00"), Format((rs!nMovImporte / rs!nMovCant) * gnIGV, "#.00"), 0, 0, Format(CDate(rs!dDocFecha), gsFormatoFecha), "", "", 60, 0, rs!cBSDescripcion, 1, 0, 1, "", "", rs!cPersCod, rs!cCtaContCod, rs!Factura, True
                lnIGVAcum = lnIGVAcum + Format(rs!nMovImporte / rs!nMovCant, "#.00")
                lnValorAcum = lnValorAcum + Format((rs!nMovImporte / rs!nMovCant) * gnIGV, "#.00")
            Else
                If rsSerie.EOF And rsSerie.BOF Then
                    lsSerie = Trim(Str(Year(gdFecSis))) & "-" & Format(CLng(oALmacen.GetCorrelaSerie(rs!cBSCod)) + 1, "00000000")
                
                    lsProveedor = oALmacen.GetProveedorSerie(lsSerie, rs!cBSCod)
                    lsFactura = oALmacen.GetProveedorFactura(lsSerie, rs!cBSCod)
                    
                    If rs!nValor = 0 Then
                        oALmacen.InsertMovBsControl Year(gdFecSis), rs!nMovNro, rs!cBSCod, lsSerie, Format(rs!nMovImporte / rs!nMovCant, "#.00"), Format((rs!nMovImporte / rs!nMovCant) * gnIGV, "#.00"), 0, 0, Format(CDate(rs!dDocFecha), gsFormatoFecha), "", "", 60, 0, rs!cBSDescripcion, 1, 0, 1, "", "", rs!cPersCod, rs!cCtaContCod, rs!Factura, True
                        lnIGVAcum = lnIGVAcum + Format(rs!nMovImporte / rs!nMovCant, "#.00")
                        lnValorAcum = lnValorAcum + Format((rs!nMovImporte / rs!nMovCant) * gnIGV, "#.00")
                    Else
                        oALmacen.InsertMovBsControl Year(gdFecSis), rs!nMovNro, rs!cBSCod, lsSerie, rsSerie!nValor, rsSerie!nIGV, 0, 0, Format(CDate(rs!dDocFecha), gsFormatoFecha), "", "", 60, 0, rs!cBSDescripcion, 1, 0, 1, "", "", rs!cPersCod, rs!cCtaContCod, rs!Factura, True
                        lnIGVAcum = lnIGVAcum + Format(rs!nMovImporte / rs!nMovCant, "#.00")
                        lnValorAcum = lnValorAcum + Format((rs!nMovImporte / rs!nMovCant) * gnIGV, "#.00")
                    End If
                Else
                    lsSerie = rsSerie!cSerie
                    
                    lsProveedor = oALmacen.GetProveedorSerie(lsSerie, rs!cBSCod)
                    lsFactura = oALmacen.GetProveedorFactura(lsSerie, rs!cBSCod)
                    
                    If rsSerie!nValor = 0 Then
                        oALmacen.InsertMovBsControl Year(gdFecSis), rs!nMovNro, rs!cBSCod, lsSerie, Format(rs!nMovImporte / rs!nMovCant, "#.00"), Format((rs!nMovImporte / rs!nMovCant) * gnIGV, "#.00"), 0, 0, Format(CDate(rs!dDocFecha), gsFormatoFecha), "", "", 60, 0, rs!cBSDescripcion, 1, 0, 1, "", "", rs!cPersCod, rs!cCtaContCod, rs!Factura, True
                        lnIGVAcum = lnIGVAcum + Format(rs!nMovImporte / rs!nMovCant, "#.00")
                        lnValorAcum = lnValorAcum + Format((rs!nMovImporte / rs!nMovCant) * gnIGV, "#.00")
                    Else
                        oALmacen.InsertMovBsControl Year(gdFecSis), rs!nMovNro, rs!cBSCod, lsSerie, rsSerie!nValor, rsSerie!nIGV, 0, 0, Format(CDate(rs!dDocFecha), gsFormatoFecha), "", "", 60, 0, rs!cBSDescripcion, 1, 0, 1, "", "", rs!cPersCod, rs!cCtaContCod, rs!Factura, True
                        lnIGVAcum = lnIGVAcum + Format(rs!nMovImporte / rs!nMovCant, "#.00")
                        lnValorAcum = lnValorAcum + Format((rs!nMovImporte / rs!nMovCant) * gnIGV, "#.00")
                    End If
                End If
                
                If Not (rsSerie.EOF And rsSerie.BOF) Then rsSerie.MoveNext
            End If
        Next lnI
        rs.MoveNext
        Prg.value = Prg.value + 1
    Wend
    
    
End Sub

Private Sub cmdProcesarS_Click()
    If Not IsDate(Me.mskFecIni.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        mskFecIni.SetFocus
        Exit Sub
    ElseIf Not IsDate(Me.mskFecFin.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        mskFecFin.SetFocus
        Exit Sub
    End If
    
    Salidas
End Sub

Private Sub cmdReporte_Click()
    Dim lnItem As Long
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    
    Dim lsArchivoN  As String
    Dim lbLibroOpen As Boolean
    
    
    Dim lsCodigo As String
    Dim lsSerie As String
    Dim lsDescripcion As String
    Dim lsAreaCod As String
    Dim lsAreadescripcion As String
    Dim lsAgeCod As String
    Dim lsAgencia As String
    Dim lsPrecio As String
    Dim lsFecha As String
    Dim lsProveedor As String
    Dim lsCtaCont As String
    Dim lsDocumento As String
    
    lbIngreso = True
    
    If Not IsDate(Me.mskFecIni.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        mskFecIni.SetFocus
        Exit Sub
    ElseIf Not IsDate(Me.mskFecFin.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        mskFecFin.SetFocus
        Exit Sub
    End If
    
    Set rs = oALmacen.GetAlmacenDetAF(CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text), True)
            
    Prg.Max = rs.RecordCount + 2
    Prg.value = Prg.Min
        
    flex.Rows = 1
    flex.Cols = 12
    
    flex.TextMatrix(0, 0) = "Codigo"
    flex.TextMatrix(0, 1) = "Serie"
    flex.TextMatrix(0, 2) = "Descripcion"
    flex.TextMatrix(0, 3) = "Area"
    flex.TextMatrix(0, 4) = "Desc.Area"
    flex.TextMatrix(0, 5) = "Age."
    flex.TextMatrix(0, 6) = "Desc.Age."
    flex.TextMatrix(0, 7) = "Precio"
    flex.TextMatrix(0, 8) = "Fecha"
    flex.TextMatrix(0, 9) = "Proeveedor"
    flex.TextMatrix(0, 10) = "CtaCont"
    flex.TextMatrix(0, 11) = "Documento"
    
   
    While Not rs.EOF
        flex.Rows = flex.Rows + 1
        
        lsCodigo = rs!cBSCod
        lsSerie = rs!cSerie
        lsDescripcion = rs!cDescripcion
        lsAreaCod = rs!cAreaCod & ""
        lsAreadescripcion = rs!cAreaDescripcion & ""
        lsAgeCod = rs!cAgeCod & ""
        lsAgencia = rs!Agencia & ""
        lsPrecio = Format(rs!nBSValor, "#,##0.00")
        lsFecha = Format(rs!dActivacion, ": " & gsFormatoFechaView)
        lsProveedor = rs!Proveedor & ""
        lsCtaCont = rs!cCtaCont
        lsDocumento = rs!cFactura
       
        flex.TextMatrix(flex.Rows - 1, 0) = lsCodigo
        flex.TextMatrix(flex.Rows - 1, 1) = lsSerie
        flex.TextMatrix(flex.Rows - 1, 2) = lsDescripcion
        flex.TextMatrix(flex.Rows - 1, 3) = lsAreaCod
        flex.TextMatrix(flex.Rows - 1, 4) = lsAreadescripcion
        flex.TextMatrix(flex.Rows - 1, 5) = lsAgeCod
        flex.TextMatrix(flex.Rows - 1, 6) = lsAgencia
        flex.TextMatrix(flex.Rows - 1, 7) = lsPrecio
        flex.TextMatrix(flex.Rows - 1, 8) = lsFecha
        flex.TextMatrix(flex.Rows - 1, 9) = lsProveedor
        flex.TextMatrix(flex.Rows - 1, 10) = lsCtaCont
        flex.TextMatrix(flex.Rows - 1, 11) = lsDocumento

        rs.MoveNext
        Prg.value = Prg.value + 1
    Wend
        
    Prg.value = Prg.Max
    
    lsArchivoN = App.path & "\Spooler\" & Format(CDate(Me.mskFecFin.Text), "yyyymmdd") & ".xls"
    
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
        If (gsOpeCod = 591507) Then
        lsPalabras = "Ingresos de Activo y BND"
        End If
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Se Imprimio el Reporte " & lsPalabras & " del " & mskFecIni.Text & " al " & mskFecFin.Text
        Set objPista = Nothing
        '**************
End Sub

Private Sub cmdReportes_Click()
    Dim lnItem As Long
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    
    Dim lsArchivoN  As String
    Dim lbLibroOpen As Boolean
    
    
    Dim lsCodigo As String
    Dim lsSerie As String
    Dim lsDescripcion As String
    Dim lsAreaCod As String
    Dim lsAreadescripcion As String
    Dim lsAgeCod As String
    Dim lsAgencia As String
    Dim lsPrecio As String
    Dim lsFecha As String
    Dim lsProveedor As String
    Dim lsCtaCont As String
    Dim lsDocumento As String
    
    lbIngreso = False
    
    If Not IsDate(Me.mskFecIni.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        mskFecIni.SetFocus
        Exit Sub
    ElseIf Not IsDate(Me.mskFecFin.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        mskFecFin.SetFocus
        Exit Sub
    End If
    
    Set rs = oALmacen.GetAlmacenDetAF(CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text), False)
            
    Prg.Max = rs.RecordCount + 2
    Prg.value = Prg.Min
        
    flex.Rows = 1
    flex.Cols = 12
    
    flex.TextMatrix(0, 0) = "Codigo"
    flex.TextMatrix(0, 1) = "Serie"
    flex.TextMatrix(0, 2) = "Descripcion"
    flex.TextMatrix(0, 3) = "Area"
    flex.TextMatrix(0, 4) = "Desc.Area"
    flex.TextMatrix(0, 5) = "Age."
    flex.TextMatrix(0, 6) = "Desc.Age."
    flex.TextMatrix(0, 7) = "Precio"
    flex.TextMatrix(0, 8) = "Fecha"
    flex.TextMatrix(0, 9) = "Persona"
    flex.TextMatrix(0, 10) = "CtaCont"
    flex.TextMatrix(0, 11) = "Documento"
    
   
    While Not rs.EOF
        flex.Rows = flex.Rows + 1
        
        lsCodigo = rs!cBSCod
        lsSerie = rs!cSerie
        lsDescripcion = rs!cDescripcion
        lsAreaCod = rs!cAreaCod & ""
        lsAreadescripcion = rs!cAreaDescripcion & ""
        lsAgeCod = rs!cAgeCod & ""
        lsAgencia = rs!Agencia & ""
        lsPrecio = Format(rs!nBSValor, "#,##0.00")
        lsFecha = Format(rs!dActivacion, ": " & gsFormatoFechaView)
        lsProveedor = rs!Proveedor & ""
        lsCtaCont = rs!cCtaCont
        lsDocumento = rs!cFactura
       
        flex.TextMatrix(flex.Rows - 1, 0) = lsCodigo
        flex.TextMatrix(flex.Rows - 1, 1) = lsSerie
        flex.TextMatrix(flex.Rows - 1, 2) = lsDescripcion
        flex.TextMatrix(flex.Rows - 1, 3) = lsAreaCod
        flex.TextMatrix(flex.Rows - 1, 4) = lsAreadescripcion
        flex.TextMatrix(flex.Rows - 1, 5) = lsAgeCod
        flex.TextMatrix(flex.Rows - 1, 6) = lsAgencia
        flex.TextMatrix(flex.Rows - 1, 7) = lsPrecio
        flex.TextMatrix(flex.Rows - 1, 8) = lsFecha
        flex.TextMatrix(flex.Rows - 1, 9) = lsProveedor
        flex.TextMatrix(flex.Rows - 1, 10) = lsCtaCont
        flex.TextMatrix(flex.Rows - 1, 11) = lsDocumento

        rs.MoveNext
        Prg.value = Prg.value + 1
    Wend
        
    Prg.value = Prg.Max
    
    lsArchivoN = App.path & "\Spooler\" & Format(CDate(Me.mskFecFin.Text), "yyyymmdd") & ".xls"
    
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
        If (gsOpeCod = 591508) Then
        lsPalabras = "Salidas de Activo y BND"
        End If
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Se Imprimio el Reporte " & lsPalabras & " del " & mskFecIni.Text & " al " & mskFecFin.Text
        Set objPista = Nothing
        '**************

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdTrasferirActivoFijo_Click()
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    
    If Not IsDate(Me.mskFecIni.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        mskFecIni.SetFocus
        Exit Sub
    ElseIf Not IsDate(Me.mskFecFin.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        mskFecFin.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Desea procesar transferir al Activo Fijo ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso") = vbNo Then Exit Sub
    
    oALmacen.TransferenciaActivoFijo Year(CDate(Me.mskFecFin.Text)), CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text)

    MsgBox "Trandferencia concluida con exito.", vbInformation, "Aviso"
        'ARLO 20160126 ***
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Se Tranferio Activo Fijo del " & mskFecIni.Text & " al " & mskFecFin.Text
        Set objPista = Nothing
        '**************
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.mskFecIni.SetFocus
End Sub

Private Sub Form_Load()
    Dim oDoc As DOperaciones
    Set oDoc = New DOperaciones
    
    Me.txtAlmacen.rs = oDoc.GetAlmacenes
    
    Me.txtAlmacen.Text = "1"
    Me.lblAmacenG.Caption = txtAlmacen.psDescripcion
    Caption = lsCaption
    
    If lbIngresoVista And Not lbSalidaVista Then
        Me.cmdProcesarS.Visible = False
        Me.cmdProcesarS.Enabled = False
        Me.cmdReporteS.Visible = False
        Me.cmdReporteS.Enabled = False
    
        Me.cmdProcesar.Visible = True
        Me.cmdProcesar.Enabled = True
        Me.cmdReporte.Visible = True
        Me.cmdReporte.Enabled = True
        
        cmdTrasferirActivoFijo.Enabled = False
        cmdTrasferirActivoFijo.Visible = False
        Me.cmdMantSalidas.Visible = False
        Me.cmdMantSalidas.Enabled = False
    ElseIf Not lbIngresoVista And lbSalidaVista Then
        Me.cmdProcesarS.Visible = True
        Me.cmdProcesarS.Enabled = True
        Me.cmdReporteS.Visible = True
        Me.cmdReporteS.Enabled = True
    
        Me.cmdProcesar.Visible = False
        Me.cmdProcesar.Enabled = False
        Me.cmdReporte.Visible = False
        Me.cmdReporte.Enabled = False
        
        cmdTrasferirActivoFijo.Enabled = False
        cmdTrasferirActivoFijo.Visible = False
        Me.cmdMantSalidas.Visible = False
        Me.cmdMantSalidas.Enabled = False
    Else
        Me.cmdProcesarS.Visible = False
        Me.cmdProcesarS.Enabled = False
        Me.cmdReporteS.Visible = False
        Me.cmdReporteS.Enabled = False
        
        Me.cmdProcesar.Visible = False
        Me.cmdProcesar.Enabled = False
        Me.cmdReporte.Visible = False
        Me.cmdReporte.Enabled = False
        
        cmdTrasferirActivoFijo.Enabled = True
        cmdTrasferirActivoFijo.Visible = True
        Me.cmdMantSalidas.Visible = True
        Me.cmdMantSalidas.Enabled = True
    End If
End Sub

Private Sub mskFecFin_GotFocus()
    mskFecFin.SelStart = 0
    mskFecFin.SelLength = 50
End Sub

Private Sub mskFecFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdProcesar.Visible Then
            Me.cmdProcesar.SetFocus
        ElseIf cmdProcesarS.Visible Then
            Me.cmdProcesarS.SetFocus
        Else
            Me.cmdTrasferirActivoFijo.SetFocus
        End If
    End If
End Sub

Private Sub mskFecIni_GotFocus()
    mskFecIni.SelStart = 0
    mskFecIni.SelLength = 50
End Sub

Private Sub mskFecIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mskFecFin.SetFocus
    End If
End Sub


Private Sub Salidas()
    Dim lnItem As Long
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim lsCadena As String
    Dim lnPagina As Long
    
    Dim lsFecha As String
    Dim lsDocumento As String
    Dim lsDescripcion As String
    Dim lsBienes As String
    Dim lsCantidad As String
    Dim lsImporte As String
    Dim lsPersona As String
    
    Dim lsDocAnt As String
    Dim lnAcumulador As Currency
        
    Dim rsSerie As ADODB.Recordset
    Set rsSerie = New ADODB.Recordset
    
    Dim lnI As Integer
    
    Dim lsArchivoN  As String
    Dim lbLibroOpen As Boolean
    
    Dim lsSerie As String
    Dim lsProveedor As String
    Dim lsFactura As String
    
    Dim lbBan As Boolean
    
    Dim lnIGVAcum As Currency
    Dim lnValorAcum As Currency
    
    If MsgBox("Desea procesar ?", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso") = vbNo Then Exit Sub
    
    Set rs = oALmacen.GetSalidasAlmacenAF(CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text), Me.txtAlmacen.Text)
    oALmacen.EliminaMovBsControl CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text), False
    
    Prg.Max = rs.RecordCount + 2
    Prg.value = Prg.Min
        
    While Not rs.EOF
        If rs!bSerie Then
            Set rsSerie = oALmacen.GetSerieMov(rs!cBSCod, rs!cDocNro, TpoDocAlmacenGuiadeRemision, rs!nMovItem)
            lbBan = True
        End If
        
        For lnI = 1 To rs!nMovCant
            If rs!bSerie = 0 Or (rsSerie.EOF And rs.RecordCount > rs.Bookmark) Then
                lsSerie = Trim(Str(Year(CDate(Me.mskFecIni.Text)))) & "-" & Format(CLng(oALmacen.GetCorrelaSerieSal(rs!cBSCod)) + 1, "00000000")
                
                If Not oALmacen.ValidaMovBsControl(Year(CDate(Me.mskFecIni.Text)), rs!nMovNro, rs!cBSCod, lsSerie, False) Then
                    oALmacen.InsertMovBsControl Year(CDate(Me.mskFecIni.Text)), rs!nMovNro, rs!cBSCod, lsSerie, Format(rs!nMovImporte / rs!nMovCant, "#.00"), Format((rs!nMovImporte / rs!nMovCant) * gnIGV, "#.00"), 0, 0, Format(CDate(rs!dDocFecha), gsFormatoFecha), rs!cAreaCod, rs!cAgeCod, rs!nPorDeprecia, 0, rs!cBSDescripcion, 1, 0, 1, "", "", "", rs!cCtaContCod, "", False
                Else
                    oALmacen.InsertMovBsControl Year(CDate(Me.mskFecIni.Text)), rs!nMovNro, rs!cBSCod, lsSerie, Format(rs!nMovImporte / rs!nMovCant, "#.00"), Format((rs!nMovImporte / rs!nMovCant) * gnIGV, "#.00"), 0, 0, Format(CDate(rs!dDocFecha), gsFormatoFecha), rs!cAreaCod, rs!cAgeCod, rs!nPorDeprecia, 0, rs!cBSDescripcion, 1, 0, 1, "", "", "", rs!cCtaContCod, "", False
                End If
                
                lnIGVAcum = lnIGVAcum + Format(rs!nMovImporte / rs!nMovCant, "#.00")
                lnValorAcum = lnValorAcum + Format((rs!nMovImporte / rs!nMovCant) * gnIGV, "#.00")
            Else
                
                lsSerie = rsSerie!cSerie
                lsProveedor = oALmacen.GetProveedorSerie(lsSerie, rs!cBSCod)
                lsFactura = oALmacen.GetProveedorFactura(lsSerie, rs!cBSCod)
                
                If Not oALmacen.ValidaMovBsControl(Year(CDate(Me.mskFecIni.Text)), rs!nMovNro, rs!cBSCod, lsSerie, False) Then
                    If rsSerie!nValor = 0 Then
                        oALmacen.InsertMovBsControl Year(CDate(Me.mskFecIni.Text)), rs!nMovNro, rs!cBSCod, lsSerie, Format(rs!nMovImporte / rs!nMovCant, "#.00"), Format((rs!nMovImporte / rs!nMovCant) * gnIGV, "#.00"), 0, 0, Format(CDate(rs!dDocFecha), gsFormatoFecha), rs!cAreaCod, rs!cAgeCod, rs!nPorDeprecia, 0, rs!cBSDescripcion, 1, 0, 1, "", "", lsProveedor, rs!cCtaContCod, lsFactura, False
                    Else
                        oALmacen.InsertMovBsControl Year(CDate(Me.mskFecIni.Text)), rs!nMovNro, rs!cBSCod, lsSerie, rsSerie!nValor, rsSerie!nIGV, 0, 0, Format(CDate(rs!dDocFecha), gsFormatoFecha), rs!cAreaCod, rs!cAgeCod, rs!nPorDeprecia, 0, rs!cBSDescripcion, 1, 0, 1, "", "", lsProveedor, rs!cCtaContCod, lsFactura, False
                    End If
                Else
                    If rsSerie!nValor = 0 Then
                        oALmacen.InsertMovBsControl Year(CDate(Me.mskFecIni.Text)), rs!nMovNro, rs!cBSCod, lsSerie, Format(rs!nMovImporte / rs!nMovCant, "#.00"), Format((rs!nMovImporte / rs!nMovCant) * gnIGV, "#.00"), 0, 0, Format(CDate(rs!dDocFecha), gsFormatoFecha), rs!cAreaCod, rs!cAgeCod, rs!nPorDeprecia, 0, rs!cBSDescripcion, 1, 0, 1, "", "", lsProveedor, rs!cCtaContCod, lsFactura, False
                    Else
                        oALmacen.InsertMovBsControl Year(CDate(Me.mskFecIni.Text)), rs!nMovNro, rs!cBSCod, lsSerie, rsSerie!nValor, rsSerie!nIGV, 0, 0, Format(CDate(rs!dDocFecha), gsFormatoFecha), rs!cAreaCod, rs!cAgeCod, rs!nPorDeprecia, 0, rs!cBSDescripcion, 1, 0, 1, "", "", lsProveedor, rs!cCtaContCod, lsFactura, False
                    End If
                End If
                rsSerie.MoveNext
            End If
        Next lnI
        rs.MoveNext
        Prg.value = Prg.value + 1
    Wend
End Sub

Private Sub GeneraReporte()
    Dim i As Integer
    Dim K As Integer
    Dim j As Integer
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
 
    For i = 0 To Me.flex.Rows - 1
        lnAcum = 0
        For j = 0 To Me.flex.Cols - 1
            xlHoja1.Cells(i + 1, j + 1) = Me.flex.TextMatrix(i, j)
            If i > 1 And j > 1 Then
                
                If IsNumeric(Me.flex.TextMatrix(i, j)) Then
                    lnAcum = lnAcum + CCur(Me.flex.TextMatrix(i, j))
                End If
            End If
        Next j
        
        If Me.flex.TextMatrix(i, 0) <> "" And i > 0 Then
            xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(i + 0))).Select
        
            xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(i + 0))).Borders(xlDiagonalDown).LineStyle = xlNone
            xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(i + 0))).Borders(xlDiagonalUp).LineStyle = xlNone
            With xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(i + 0))).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(i + 0))).Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(i + 0))).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(i + 0))).Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(i + 0))).Borders(xlInsideVertical).LineStyle = xlNone
            
            lnFilaMarcaIni = i + 1
        End If
        
        If i > 1 Then
            'xlHoja1.Range("A1:K" & Trim(Str(Me.flex.Rows))).Select
            'lnFilaMarcaIni = 0
            
            'VSQL = Format(lnAcum, "#,##0.00")  ' "=SUMA(" & Trim(ExcelColumnaString(3)) & Trim(I + 1) & ":" & Trim(ExcelColumnaString(Me.Flex.Cols)) & Trim(I + 1) & ")"
            'xlHoja1.Cells(I + 1, Me.Flex.Cols + 1).Formula = VSQL
            'xlHoja1.Cells(I + 1, Me.flex.Cols + 1) = VSQL
        End If
    Next i
        
    xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(i + 0))).Select

    xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(i + 0))).Borders(xlDiagonalDown).LineStyle = xlNone
    xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(i + 0))).Borders(xlDiagonalUp).LineStyle = xlNone
    With xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(i + 0))).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(i + 0))).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(i + 0))).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(i + 0))).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    xlHoja1.Range("A" & lnFilaMarcaIni & ":K" & Trim(Str(i + 0))).Borders(xlInsideVertical).LineStyle = xlNone
    
    lnFilaMarcaIni = i + 1
        
    xlHoja1.Range("A1:A" & Trim(Str(Me.flex.Rows))).Font.Bold = True
    xlHoja1.Range("B1:B" & Trim(Str(Me.flex.Rows))).Font.Bold = True
    xlHoja1.Range("G1:G" & Trim(Str(Me.flex.Rows))).Font.Bold = True
    xlHoja1.Range("1:1").Font.Bold = True

    xlHoja1.Range("E2:G" & Trim(Str(Me.flex.Rows))).NumberFormat = "#,##0.00"

    xlHoja1.Cells.Select
    xlHoja1.Cells.EntireColumn.AutoFit



'************************************

   With xlHoja1.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    xlHoja1.PageSetup.PrintArea = ""
    With xlHoja1.PageSetup
        .LeftHeader = ""
        .CenterHeader = "&""Arial,Negrita""&18Listado de " & IIf(lbIngreso, "Ingresos", "Salidas") & " " & Format(CDate(Me.mskFecIni.Text), "mmmm yyyy")
        .RightHeader = "&P"
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
'        .LeftMargin = Application.InchesToPoints(0)
'        .RightMargin = Application.InchesToPoints(0)
'        .TopMargin = Application.InchesToPoints(0.39)
'        .BottomMargin = Application.InchesToPoints(0)
'        .HeaderMargin = Application.InchesToPoints(0.14)
'        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        '.PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 60
    End With
       
    xlHoja1.Columns("H:K").Select
    With xlHoja1.Range("H:K")
        .HorizontalAlignment = xlJustify
        .VerticalAlignment = xlJustify
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    
    
    xlHoja1.Columns("A:K").Select
    xlHoja1.Range("A:K").Subtotal GroupBy:=3, Function:=xlSum, TotalList:=Array(8), Replace:=True, PageBreaks:=False, SummaryBelowData:=True
    xlHoja1.Range("A:K").Subtotal GroupBy:=1, Function:=xlCount, TotalList:=Array(1), Replace:=False, PageBreaks:=False, SummaryBelowData:=True
   
    xlHoja1.Columns("I:I").Select
    xlHoja1.Range("I:I").NumberFormat = "#,##0.00"
    With xlHoja1.Range("I:I")
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlJustify
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    
    xlHoja1.Select
    With xlHoja1.Cells.Font
        .Name = "Arial"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
End Sub
