VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAdeudRepVenc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adeudados: Pagares x Fecha de Vencimiento"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   Icon            =   "frmAdeudRepVenc.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraopciones 
      Caption         =   "Banco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1125
      Left            =   90
      TabIndex        =   6
      Top             =   630
      Width           =   4785
      Begin VB.CheckBox chktodos 
         Caption         =   "&Todos"
         Height          =   270
         Left            =   3810
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Checked
         Width           =   765
      End
      Begin Sicmact.TxtBuscar txtCodObjeto 
         Height          =   345
         Left            =   1050
         TabIndex        =   7
         Top             =   180
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Enabled         =   0   'False
         sTitulo         =   ""
         EnabledText     =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Objeto :"
         Height          =   255
         Left            =   105
         TabIndex        =   11
         Top             =   270
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion :"
         Height          =   195
         Left            =   105
         TabIndex        =   10
         Top             =   630
         Width           =   930
      End
      Begin VB.Label lblObjDesc 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1050
         TabIndex        =   9
         Top             =   570
         Width           =   3570
      End
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      Height          =   360
      Left            =   1980
      TabIndex        =   5
      Top             =   1860
      Width           =   1440
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   3420
      TabIndex        =   4
      Top             =   1860
      Width           =   1440
   End
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   4785
      Begin MSMask.MaskEdBox txtfechaDel 
         Height          =   330
         Left            =   1800
         TabIndex        =   1
         Top             =   150
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaAl 
         Height          =   330
         Left            =   3465
         TabIndex        =   2
         Top             =   150
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Al:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   3075
         TabIndex        =   12
         Top             =   195
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Vcmto. del:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   210
         Width           =   1560
      End
   End
   Begin MSComctlLib.ProgressBar Barra 
      Height          =   165
      Left            =   2565
      TabIndex        =   13
      Top             =   2340
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar Estado 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   14
      Top             =   2280
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAdeudRepVenc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsObjetos() As String
Dim lbExcel As Boolean
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Dim lsArchivo As String
Dim n As Integer
Dim lbBancos As Boolean
Dim lbCortoPlazo As Boolean
Dim lbLoad As Boolean
Dim dbCmact As DConecta
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Sub chkTodos_Click()
    If Me.chktodos.value = 1 Then
        Me.txtCodObjeto.Enabled = False
        Me.lblObjDesc = ""
        Me.txtCodObjeto = ""
    Else
        Me.txtCodObjeto.Enabled = True
    End If
End Sub
Public Sub Inicio(Optional pbCortoPlazo As Boolean = False)
    lbCortoPlazo = pbCortoPlazo
    Me.Show 1
End Sub
Private Sub chkTodos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmdGenerar.SetFocus
    End If
End Sub

Private Sub cmdGenerar_Click()
  On Error GoTo ErrorGenerar

    If chktodos.value = 0 Then
        If txtCodObjeto = "" Then
            MsgBox "No se selecciono Cuenta de Adeudado", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    If ValFecha(txtfechaDel) = False Then
        Exit Sub
    End If

    If ValFecha(txtFechaAl) = False Then
        Exit Sub
    End If

    lbExcel = False
 
    n = 0
    Call DatosReporteGeneral(Trim(txtCodObjeto))
     
    Exit Sub
ErrorGenerar:
    MsgBox "Error N?[" & Err.Number & "] " & Err.Description, vbInformation, "Aviso"
    If lbExcel = True Then
        ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    If lbLoad = False Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim n As Long, m As Long

    Me.Caption = gsOpeDesc
    CentraForm Me
    lbLoad = True
    Set dbCmact = New DConecta
    dbCmact.AbreConexion
    txtfechaDel = gdFecSis
    txtFechaAl = gdFecSis
    ReDim lsObjetos(4, 0)
    n = 0
    sql = "Select * from OpeObj where cOpeCod ='" & gsOpeCod & "' and cOpeObjOrden = '0'"
    Set rs = dbCmact.CargaRecordSet(sql)
    If Not RSVacio(rs) Then
        Do While Not rs.EOF
            n = n + 1
            ReDim Preserve lsObjetos(4, n)
            lsObjetos(1, n) = Trim(rs!cObjetoCod)
            lsObjetos(2, n) = Trim(rs!nOpeObjNiv)
            lsObjetos(3, n) = Trim(rs!cOpeObjFiltro)
            lsObjetos(4, n) = Trim(rs!cOpeObjOrden)
            rs.MoveNext
        Loop
    Else
        RSClose rs
        MsgBox "No se han Definido Objetos para Reporte", vbInformation, "Aviso"
        lbLoad = False
        Exit Sub
    End If
    RSClose rs
    
    Dim oIF As New DCajaCtasIF
    Me.txtCodObjeto.rs = oIF.CargaCtasIF(Mid(gsOpeCod, 3, 1), lsObjetos(3, 1), MuestraInstituciones)
    Set oIF = Nothing
    
End Sub

Private Sub DatosReporteGeneral(lsBanco As String)
    Dim rs As New ADODB.Recordset
    Dim lnTotal As Integer, j As Integer
    Dim lnIndiceVac As Double
    Dim sArchGrabar As String
    Dim lbLibroOpen As Boolean
    Dim fs As New Scripting.FileSystemObject
    Dim lbExisteHoja As Boolean
    Dim nfil As Integer
    Dim nFilTemp0 As Integer
    Dim nFilTemp1(1 To 6) As String 'Total
    Dim nFilTemp2(1 To 6) As String 'Por Fecha
    Dim nFilTemp3(1 To 6) As String ' Por Entidadd
    Dim nFilTemp4(1 To 6) As String 'x Fuente
    Dim nCant As Integer
    Dim i As Integer
    
    Dim nTempFecha As String
    Dim nTempEntidad As String
    Dim nTempFte As String
    
On Error GoTo ErrBegin
    
    sArchGrabar = App.path & "\Spooler\PagaresxVenc" & Format(txtfechaDel, "ddMMYYYY") & "_" & Format(txtFechaAl, "ddMMYYYY") & ".xlsx"
    
    Me.Barra.value = 0
    Me.Estado.Panels(1).Text = ""
     
    Dim oIF As New NCajaAdeudados
    Dim oDAdeud As DCaja_Adeudados
    Screen.MousePointer = vbHourglass
    Set oDAdeud = New DCaja_Adeudados
    lnIndiceVac = oDAdeud.CargaIndiceVAC(Format(txtFechaAl.Text, "MM/dd/YYYY")) ' CDate("01/" & Format(Month(gdFecSis), "00") & "/" & Format(Year(gdFecSis), "0000")) - 1)
    'lnIndiceVac = oDAdeud.CargaIndiceVAC(CDate("01/" & Format(Month(gdFecSis), "00") & "/" & Format(Year(gdFecSis), "0000")) - 1)
    
    Set oDAdeud = Nothing
    
    Set rs = CargaDatosPagaresxFecha(lsBanco, lsObjetos(3, 1), , lnIndiceVac, txtfechaDel, txtFechaAl, gsOpeCod)
    lnTotal = rs.RecordCount
    If Not RSVacio(rs) Then
     
        Set xlAplicacion = New Excel.Application
        lbLibroOpen = ExcelBegin(sArchGrabar, xlAplicacion, xlLibro)
         
        If lbLibroOpen Then
           Set xlHoja1 = xlLibro.Worksheets(1)
           
           xlHoja1.Name = "PAGARES X VCTO"
           nfil = 1
           xlHoja1.Cells(nfil, 1) = gsNomCmac
           xlHoja1.Cells(nfil, 6) = "MONEDA " & IIf(Mid(gsOpeCod, 3, 1) = gMonedaNacional, "NACIONAL", "EXTRANJERA")
           xlHoja1.Range("A" & Trim(Str(nfil)) & ":E" & Trim(Str(nfil))).MergeCells = True
           xlHoja1.Range("F" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).MergeCells = True
            
           xlHoja1.Cells(nfil, 1).HorizontalAlignment = xlLeft
           xlHoja1.Cells(nfil, 6).HorizontalAlignment = xlCenter
           xlHoja1.Cells(nfil, 1).Font.Bold = True
           xlHoja1.Cells(nfil, 6).Font.Bold = True
           
           nfil = nfil + 2
           xlHoja1.Cells(nfil, 1) = "LISTADO DE PAGARES POR FECHA DE VENCIMIENTO"
           xlHoja1.Range("A" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).MergeCells = True
           xlHoja1.Cells(nfil, 1).HorizontalAlignment = xlCenter
           xlHoja1.Cells(nfil, 1).Font.Bold = True
           xlHoja1.Cells(nfil, 1).Font.Underline = True
           nfil = nfil + 1
           xlHoja1.Cells(nfil, 1) = "Periodo del " & txtfechaDel.Text & " Al " & txtFechaAl.Text
           xlHoja1.Range("A" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).MergeCells = True
           xlHoja1.Cells(nfil, 1).HorizontalAlignment = xlCenter
           xlHoja1.Cells(nfil, 1).Font.Bold = True
           nfil = nfil + 1
            
           'nTempFecha = rs!dFechaVenc
           
           Do While Not rs.EOF
 
                If nTempFecha <> rs!dFechaVenc Then
                    If nfil > 5 Then
                    
                        '**************************************
                        If Len(Trim(nTempFte)) > 0 Then
                            nfil = nfil + 1
                            xlHoja1.Cells(nfil, 1) = "Total x Linea"
                            xlHoja1.Cells(nfil, 1).Font.Bold = True
                            
                            xlHoja1.Range("C" & Trim(Str(nfil)) & ":C" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp4(1))
                            xlHoja1.Range("D" & Trim(Str(nfil)) & ":D" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp4(2))
                            xlHoja1.Range("E" & Trim(Str(nfil)) & ":E" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp4(3))
                            xlHoja1.Range("F" & Trim(Str(nfil)) & ":F" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp4(4))
                            xlHoja1.Range("G" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp4(5))
                            xlHoja1.Range("H" & Trim(Str(nfil)) & ":H" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp4(6))
                            
                            For i = 1 To 6
                                nFilTemp4(i) = ""
                            Next
                        End If
                        '**************************************
                    
                        nfil = nfil + 1
                        xlHoja1.Cells(nfil, 1) = "Total x Entidad"
                        xlHoja1.Cells(nfil, 1).Font.Bold = True
                        ''''''''''''''''''''''''''''''''''''''''''
                        xlHoja1.Range("C" & Trim(Str(nfil)) & ":C" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp2(1))
                        xlHoja1.Range("D" & Trim(Str(nfil)) & ":D" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp2(2))
                        xlHoja1.Range("E" & Trim(Str(nfil)) & ":E" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp2(3))
                        xlHoja1.Range("F" & Trim(Str(nfil)) & ":F" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp2(4))
                        xlHoja1.Range("G" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp2(5))
                        xlHoja1.Range("H" & Trim(Str(nfil)) & ":H" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp2(6))
                        
                        'Cabecera
                        ExcelCuadro xlHoja1, 1, nFilTemp0 - 2, 8, nFilTemp0 - 1
                        
                        'Cuadro del centro
                        ExcelCuadro xlHoja1, 1, CCur(nFilTemp0), 8, nfil - 1
                        If Len(Trim(nTempFte)) > 0 Then
                            ExcelCuadro xlHoja1, 1, nfil - 1, 8, CCur(nfil)
                        Else
                            'Cuadro del total
                            ExcelCuadro xlHoja1, 1, CCur(nfil), 8, CCur(nfil)
                        End If
                        
                        For i = 1 To 6
                            nFilTemp2(i) = ""
                        Next
                        ''''''''''''''''''''''''''''''''''''''''''
                        nfil = nfil + 2
                        xlHoja1.Cells(nfil, 1) = "Total x Fecha"
                        xlHoja1.Cells(nfil, 1).Font.Bold = True
                        ''''''''''''''''''''''''''''''''''''''''''
                        xlHoja1.Range("C" & Trim(Str(nfil)) & ":C" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp1(1))
                        xlHoja1.Range("D" & Trim(Str(nfil)) & ":D" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp1(2))
                        xlHoja1.Range("E" & Trim(Str(nfil)) & ":E" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp1(3))
                        xlHoja1.Range("F" & Trim(Str(nfil)) & ":F" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp1(4))
                        xlHoja1.Range("G" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp1(5))
                        xlHoja1.Range("H" & Trim(Str(nfil)) & ":H" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp1(6))
                        
                        'Cuadro del total
                        ExcelCuadro xlHoja1, 1, CCur(nfil), 8, CCur(nfil)
                        
                        For i = 1 To 6
                            nFilTemp1(i) = ""
                        Next
                        ''''''''''''''''''''''''''''''''''''''''''
                        nfil = nfil + 1
                    End If
                    nfil = nfil + 1
                    nTempFecha = rs!dFechaVenc
                    nTempEntidad = rs!cPersCod
                    nTempFte = rs!cCodLinCred
                    xlHoja1.Cells(nfil, 1) = "FECHA"
                    xlHoja1.Cells(nfil, 2) = "'" & Format(rs!dVencimiento, "dd/MM/YYYY")
                    xlHoja1.Range("A" & Trim(Str(nfil)) & ":H" & Trim(Str(nfil))).Font.Bold = True
                    ExcelCuadro xlHoja1, 1, CCur(nfil), 2, CCur(nfil)
                    
                    xlHoja1.Range("A" & Trim(Str(nfil)) & ":B" & Trim(Str(nfil))).Interior.ColorIndex = 36
                    xlHoja1.Range("A" & Trim(Str(nfil)) & ":A" & Trim(Str(nfil))).Font.ColorIndex = 3
                    xlHoja1.Range("B" & Trim(Str(nfil)) & ":B" & Trim(Str(nfil))).Font.ColorIndex = 5
                    
                    nfil = nfil + 1
                    
                    nfil = nfil + 1
                    nTempEntidad = rs!cPersCod
                    xlHoja1.Cells(nfil, 1) = "Entidad"
                    xlHoja1.Cells(nfil, 2) = rs!cPersNombre
                    xlHoja1.Range("B" & Trim(Str(nfil)) & ":H" & Trim(Str(nfil))).MergeCells = True
                    xlHoja1.Cells(nfil, 1).Font.Underline = True
                    
                    '**************************************
                    If Len(Trim(rs!cCodLinCred)) > 0 Then
                        nfil = nfil + 1
                        xlHoja1.Cells(nfil, 1) = "Linea"
                        xlHoja1.Cells(nfil, 2) = rs!cDesLinCred
                        xlHoja1.Range("B" & Trim(Str(nfil)) & ":H" & Trim(Str(nfil))).MergeCells = True
                    End If
                    '**************************************
                    nfil = nfil + 1
                    xlHoja1.Cells(nfil, 1) = "Numero"
                    xlHoja1.Cells(nfil, 2) = "Tasa"
                    xlHoja1.Cells(nfil, 3) = "Capital"
                    xlHoja1.Cells(nfil, 4) = "Cuota Vencida"
                    xlHoja1.Cells(nfil, 8) = "Capital"
                    xlHoja1.Range("D" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).MergeCells = True
                    nfil = nfil + 1
                    xlHoja1.Cells(nfil, 1) = "Pagar?"
                    xlHoja1.Cells(nfil, 2) = "%"
                    xlHoja1.Cells(nfil, 3) = "Actual"
                    xlHoja1.Cells(nfil, 4) = "Capital"
                    xlHoja1.Cells(nfil, 5) = "Interes"
                    xlHoja1.Cells(nfil, 6) = "Comisi?n"
                    xlHoja1.Cells(nfil, 7) = "Total"
                    xlHoja1.Cells(nfil, 8) = "Al " & txtFechaAl.Text
                
                    ExcelCuadro xlHoja1, 4, CCur(nfil), 7, CCur(nfil)
                    
                    xlHoja1.Range("A" & Trim(Str(nfil - 1)) & ":H" & Trim(Str(nfil))).HorizontalAlignment = xlCenter
                    xlHoja1.Range("A" & Trim(Str(nfil - 3)) & ":H" & Trim(Str(nfil))).Font.Bold = True
                    
                    nFilTemp0 = nfil + 1
                Else
                    If nTempEntidad <> rs!cPersCod Then
                        If nfil > 5 Then
                            
                            '**************************************
                            If Len(Trim(nTempFte)) > 0 Then
                                nfil = nfil + 1
                                xlHoja1.Cells(nfil, 1) = "Total x Linea"
                                xlHoja1.Cells(nfil, 1).Font.Bold = True
                                
                                xlHoja1.Range("C" & Trim(Str(nfil)) & ":C" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp4(1))
                                xlHoja1.Range("D" & Trim(Str(nfil)) & ":D" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp4(2))
                                xlHoja1.Range("E" & Trim(Str(nfil)) & ":E" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp4(3))
                                xlHoja1.Range("F" & Trim(Str(nfil)) & ":F" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp4(4))
                                xlHoja1.Range("G" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp4(5))
                                xlHoja1.Range("H" & Trim(Str(nfil)) & ":H" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp4(6))
                                
                                For i = 1 To 6
                                    nFilTemp4(i) = ""
                                Next
                            End If
                            '**************************************
                            
                            nfil = nfil + 1
                            xlHoja1.Cells(nfil, 1) = "Total x Entidad"
                            ''''''''''''''''''''''''''''''''''''''''''
                            xlHoja1.Range("C" & Trim(Str(nfil)) & ":C" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp2(1))
                            xlHoja1.Range("D" & Trim(Str(nfil)) & ":D" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp2(2))
                            xlHoja1.Range("E" & Trim(Str(nfil)) & ":E" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp2(3))
                            xlHoja1.Range("F" & Trim(Str(nfil)) & ":F" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp2(4))
                            xlHoja1.Range("G" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp2(5))
                            xlHoja1.Range("H" & Trim(Str(nfil)) & ":H" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp2(6))
                            
                            'Cabecera
                            ExcelCuadro xlHoja1, 1, nFilTemp0 - 2, 8, nFilTemp0 - 1
                            
                            'Cuadro del centro
                            ExcelCuadro xlHoja1, 1, CCur(nFilTemp0), 8, nfil - 1
                            'Cuadro del total
                            If Len(Trim(nTempFte)) > 0 Then
                                ExcelCuadro xlHoja1, 1, nfil - 1, 7, nfil
                            Else
                                ExcelCuadro xlHoja1, 1, nfil, 8, nfil
                            End If
                            
                            For i = 1 To 6
                                nFilTemp2(i) = ""
                            Next
                            ''''''''''''''''''''''''''''''''''''''''''
                            nfil = nfil + 1
                        End If
                        nfil = nfil + 1
                        nTempEntidad = rs!cPersCod
                        nTempFte = rs!cCodLinCred
                        xlHoja1.Cells(nfil, 1) = "Entidad"
                        xlHoja1.Cells(nfil, 2) = rs!cPersNombre
                        xlHoja1.Range("B" & Trim(Str(nfil)) & ":H" & Trim(Str(nfil))).MergeCells = True
                        xlHoja1.Cells(nfil, 1).Font.Underline = True
                        '**************************************
                        If Len(Trim(rs!cCodLinCred)) > 0 Then
                            nfil = nfil + 1
                            xlHoja1.Cells(nfil, 1) = "Linea"
                            xlHoja1.Cells(nfil, 2) = rs!cDesLinCred
                            xlHoja1.Range("B" & Trim(Str(nfil)) & ":H" & Trim(Str(nfil))).MergeCells = True
                        End If
                        '**************************************
                        nfil = nfil + 1
                        xlHoja1.Cells(nfil, 1) = "Numero"
                        xlHoja1.Cells(nfil, 2) = "Tasa"
                        xlHoja1.Cells(nfil, 3) = "Capital"
                        xlHoja1.Cells(nfil, 4) = "Cuota Vencida"
                        xlHoja1.Cells(nfil, 8) = "Capital"
                        xlHoja1.Range("D" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).MergeCells = True
                        nfil = nfil + 1
                        xlHoja1.Cells(nfil, 1) = "Pagar?"
                        xlHoja1.Cells(nfil, 2) = "%"
                        xlHoja1.Cells(nfil, 3) = "Actual"
                        xlHoja1.Cells(nfil, 4) = "Capital"
                        xlHoja1.Cells(nfil, 5) = "Interes"
                        xlHoja1.Cells(nfil, 6) = "Comisi?n"
                        xlHoja1.Cells(nfil, 7) = "Total"
                        xlHoja1.Cells(nfil, 8) = "Al " & txtFechaAl.Text
                        
                        ExcelCuadro xlHoja1, 4, nfil, 7, nfil
                        
                        nFilTemp0 = nfil + 1
                        
                        xlHoja1.Range("A" & Trim(Str(nfil - 1)) & ":H" & Trim(Str(nfil))).HorizontalAlignment = xlCenter
                        xlHoja1.Range("A" & Trim(Str(nfil - 2)) & ":H" & Trim(Str(nfil))).Font.Bold = True
                    
                    Else
                        If nTempFte <> rs!cCodLinCred And Len(Trim(rs!cCodLinCred)) > 0 Then
                            '*---********************************************
                            If nfil > 5 Then
                                
                                '**************************************
                                If Len(Trim(nTempFte)) > 0 Then
                                    nfil = nfil + 1
                                    xlHoja1.Cells(nfil, 1) = "Total x Linea"
                                    xlHoja1.Cells(nfil, 1).Font.Bold = True
                                    
                                    xlHoja1.Range("C" & Trim(Str(nfil)) & ":C" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp4(1))
                                    xlHoja1.Range("D" & Trim(Str(nfil)) & ":D" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp4(2))
                                    xlHoja1.Range("E" & Trim(Str(nfil)) & ":E" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp4(3))
                                    xlHoja1.Range("F" & Trim(Str(nfil)) & ":F" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp4(4))
                                    xlHoja1.Range("G" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp4(5))
                                    xlHoja1.Range("H" & Trim(Str(nfil)) & ":H" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp4(6))
                                    
                                    For i = 1 To 6
                                        nFilTemp4(i) = ""
                                    Next
                                     
                                    'Cabecera
                                    ExcelCuadro xlHoja1, 1, nFilTemp0 - 2, 8, nFilTemp0 - 1
                                    'Cuadro del centro
                                    ExcelCuadro xlHoja1, 1, nFilTemp0, 8, nfil - 1
                                    'Cuadro del total
                                    ExcelCuadro xlHoja1, 1, nfil, 8, nfil
                                    
                                    nfil = nfil + 1
                                End If
                            End If
                            
'                            If Len(Trim(rs!cCodLinCred)) > 0 Then
                                nfil = nfil + 1
                            
                                xlHoja1.Cells(nfil, 1) = "Linea"
                                xlHoja1.Cells(nfil, 2) = rs!cDesLinCred
                                xlHoja1.Range("B" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).MergeCells = True
                                nfil = nfil + 1
                                xlHoja1.Cells(nfil, 1) = "Numero"
                                xlHoja1.Cells(nfil, 2) = "Tasa"
                                xlHoja1.Cells(nfil, 3) = "Capital"
                                xlHoja1.Cells(nfil, 4) = "Cuota Vencida"
                                xlHoja1.Cells(nfil, 8) = "Capital"
                                xlHoja1.Range("D" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).MergeCells = True
                                nfil = nfil + 1
                                xlHoja1.Cells(nfil, 1) = "Pagar?"
                                xlHoja1.Cells(nfil, 2) = "%"
                                xlHoja1.Cells(nfil, 3) = "Actual"
                                xlHoja1.Cells(nfil, 4) = "Capital"
                                xlHoja1.Cells(nfil, 5) = "Interes"
                                xlHoja1.Cells(nfil, 6) = "Comisi?n"
                                xlHoja1.Cells(nfil, 7) = "Total"
                                xlHoja1.Cells(nfil, 8) = "Al " & txtFechaAl.Text
                                
                                ExcelCuadro xlHoja1, 4, nfil, 7, nfil
                                
                                nFilTemp0 = nfil + 1
                                
                                xlHoja1.Range("A" & Trim(Str(nfil - 1)) & ":H" & Trim(Str(nfil))).HorizontalAlignment = xlCenter
                                xlHoja1.Range("A" & Trim(Str(nfil - 2)) & ":H" & Trim(Str(nfil))).Font.Bold = True
                                '*---********************************************
'                            End If
                            nTempFte = rs!cCodLinCred
                        End If
                        
                        
                    End If
                End If
                nfil = nfil + 1
                
                nCant = nCant + 1
                
                xlHoja1.Cells(nfil, 1) = rs!cCtaIFDesc
                xlHoja1.Cells(nfil, 2) = rs!nTasaInteres
                xlHoja1.Cells(nfil, 3) = rs!nSaldoCap
                xlHoja1.Cells(nfil, 4) = rs!nCapitalCon_VAC
                xlHoja1.Cells(nfil, 5) = rs!NInteresCon_VAC
                xlHoja1.Cells(nfil, 6) = rs!nComisionCon_VAC
                'xlHoja1.Cells(nFil, 8) = rs!nSaldoCap - rs!nCapitalCon_VAC
                
                xlHoja1.Range("G" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).Formula = "=+SUM(D" & Trim(Str(nfil)) & ":F" & Trim(Str(nfil)) & ")"
                xlHoja1.Range("H" & Trim(Str(nfil)) & ":H" & Trim(Str(nfil))).Formula = "=+C" & Trim(Str(nfil)) & "-D" & Trim(Str(nfil)) & ""
                xlHoja1.Range("B" & Trim(Str(nfil)) & ":H" & Trim(Str(nfil))).NumberFormat = "#,##0.00"
                 
                nFilTemp1(1) = nFilTemp1(1) & "+C" & Trim(Str(nfil))
                nFilTemp2(1) = nFilTemp2(1) & "+C" & Trim(Str(nfil))
                nFilTemp3(1) = nFilTemp3(1) & "+C" & Trim(Str(nfil))
                If Len(Trim(rs!cCodLinCred)) > 0 Then
                    nFilTemp4(1) = nFilTemp4(1) & "+C" & Trim(Str(nfil))
                End If
                
                nFilTemp1(2) = nFilTemp1(2) & "+D" & Trim(Str(nfil))
                nFilTemp2(2) = nFilTemp2(2) & "+D" & Trim(Str(nfil))
                nFilTemp3(2) = nFilTemp3(2) & "+D" & Trim(Str(nfil))
                If Len(Trim(rs!cCodLinCred)) > 0 Then
                    nFilTemp4(2) = nFilTemp4(2) & "+D" & Trim(Str(nfil))
                End If
                
                nFilTemp1(3) = nFilTemp1(3) & "+E" & Trim(Str(nfil))
                nFilTemp2(3) = nFilTemp2(3) & "+E" & Trim(Str(nfil))
                nFilTemp3(3) = nFilTemp3(3) & "+E" & Trim(Str(nfil))
                If Len(Trim(rs!cCodLinCred)) > 0 Then
                    nFilTemp4(3) = nFilTemp4(3) & "+E" & Trim(Str(nfil))
                End If
                
                nFilTemp1(4) = nFilTemp1(4) & "+F" & Trim(Str(nfil))
                nFilTemp2(4) = nFilTemp2(4) & "+F" & Trim(Str(nfil))
                nFilTemp3(4) = nFilTemp3(4) & "+F" & Trim(Str(nfil))
                If Len(Trim(rs!cCodLinCred)) > 0 Then
                    nFilTemp4(4) = nFilTemp4(4) & "+F" & Trim(Str(nfil))
                End If
                
                nFilTemp1(5) = nFilTemp1(5) & "+G" & Trim(Str(nfil))
                nFilTemp2(5) = nFilTemp2(5) & "+G" & Trim(Str(nfil))
                nFilTemp3(5) = nFilTemp3(5) & "+G" & Trim(Str(nfil))
                If Len(Trim(rs!cCodLinCred)) > 0 Then
                    nFilTemp4(5) = nFilTemp4(5) & "+G" & Trim(Str(nfil))
                End If
                
                nFilTemp1(6) = nFilTemp1(6) & "+H" & Trim(Str(nfil))
                nFilTemp2(6) = nFilTemp2(6) & "+H" & Trim(Str(nfil))
                nFilTemp3(6) = nFilTemp3(6) & "+H" & Trim(Str(nfil))
                If Len(Trim(rs!cCodLinCred)) > 0 Then
                    nFilTemp4(6) = nFilTemp4(6) & "+H" & Trim(Str(nfil))
                End If
                
                Me.Barra.value = Int(nCant / lnTotal * 100)
                Me.Estado.Panels(1).Text = "Avance :" & Format(nCant / lnTotal * 100, "#0.00") & "%"
                DoEvents
                rs.MoveNext
           Loop
            
            '**************************************
            If Len(Trim(nTempFte)) > 0 Then
                nfil = nfil + 1
                xlHoja1.Cells(nfil, 1) = "Total x Fuente"
                xlHoja1.Cells(nfil, 1).Font.Bold = True
                ''''''''''''''''''''''''''''''''''''''''''
                xlHoja1.Range("C" & Trim(Str(nfil)) & ":C" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp4(1))
                xlHoja1.Range("D" & Trim(Str(nfil)) & ":D" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp4(2))
                xlHoja1.Range("E" & Trim(Str(nfil)) & ":E" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp4(3))
                xlHoja1.Range("F" & Trim(Str(nfil)) & ":F" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp4(4))
                xlHoja1.Range("G" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp4(5))
                xlHoja1.Range("H" & Trim(Str(nfil)) & ":H" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp4(6))
            End If
            '**************************************
            
            
            nfil = nfil + 1
            xlHoja1.Cells(nfil, 1) = "Total x Entidad"
            xlHoja1.Cells(nfil, 1).Font.Bold = True
            ''''''''''''''''''''''''''''''''''''''''''
            xlHoja1.Range("C" & Trim(Str(nfil)) & ":C" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp2(1))
            xlHoja1.Range("D" & Trim(Str(nfil)) & ":D" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp2(2))
            xlHoja1.Range("E" & Trim(Str(nfil)) & ":E" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp2(3))
            xlHoja1.Range("F" & Trim(Str(nfil)) & ":F" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp2(4))
            xlHoja1.Range("G" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp2(5))
            xlHoja1.Range("H" & Trim(Str(nfil)) & ":H" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp2(6))
            
            'Cabecera
            ExcelCuadro xlHoja1, 1, nFilTemp0 - 2, 8, nFilTemp0 - 1
            
            'Cuadro del centro
            ExcelCuadro xlHoja1, 1, nFilTemp0, 8, nfil - 1
            'Cuadro del total
            If Len(Trim(nTempFte)) > 0 Then
                ExcelCuadro xlHoja1, 1, nfil - 1, 8, nfil
            Else
                ExcelCuadro xlHoja1, 1, nfil, 8, nfil
            End If
            
            For i = 1 To 6
                nFilTemp2(i) = ""
            Next
            ''''''''''''''''''''''''''''''''''''''''''
            nfil = nfil + 2
            xlHoja1.Cells(nfil, 1) = "Total x Fecha"
            xlHoja1.Cells(nfil, 1).Font.Bold = True
            ''''''''''''''''''''''''''''''''''''''''''
            xlHoja1.Range("C" & Trim(Str(nfil)) & ":C" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp1(1))
            xlHoja1.Range("D" & Trim(Str(nfil)) & ":D" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp1(2))
            xlHoja1.Range("E" & Trim(Str(nfil)) & ":E" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp1(3))
            xlHoja1.Range("F" & Trim(Str(nfil)) & ":F" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp1(4))
            xlHoja1.Range("G" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp1(5))
            xlHoja1.Range("H" & Trim(Str(nfil)) & ":H" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp1(6))
            
            'Cuadro del total
            ExcelCuadro xlHoja1, 1, nfil, 8, nfil
                        
            For i = 1 To 6
                nFilTemp1(i) = ""
            Next
            ''''''''''''''''''''''''''''''''''''''''''
           
            nfil = nfil + 2
            xlHoja1.Cells(nfil, 1) = "Resumen"
            xlHoja1.Cells(nfil, 2) = "Elementos"
            xlHoja1.Cells(nfil, 3) = "Capital"
            xlHoja1.Cells(nfil, 4) = "Cuota Vencida"
            xlHoja1.Cells(nfil, 8) = "Capital"
            xlHoja1.Range("D" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).MergeCells = True
            nfil = nfil + 1
            xlHoja1.Cells(nfil, 3) = "Actual"
            xlHoja1.Cells(nfil, 4) = "Capital"
            xlHoja1.Cells(nfil, 5) = "Interes"
            xlHoja1.Cells(nfil, 6) = "Comisi?n"
            xlHoja1.Cells(nfil, 7) = "Total"
            xlHoja1.Cells(nfil, 8) = "Al " & txtFechaAl.Text
            
            ExcelCuadro xlHoja1, 4, nfil, 7, nfil
            
            xlHoja1.Range("A" & Trim(Str(nfil - 1)) & ":H" & Trim(Str(nfil))).HorizontalAlignment = xlCenter
            xlHoja1.Range("A" & Trim(Str(nfil - 3)) & ":H" & Trim(Str(nfil))).Font.Bold = True
            
            ''''''''''''''''''''''''''''''''''''''''''
            nfil = nfil + 1
            xlHoja1.Cells(nfil, 2) = nCant
            xlHoja1.Range("C" & Trim(Str(nfil)) & ":C" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp3(1))
            xlHoja1.Range("D" & Trim(Str(nfil)) & ":D" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp3(2))
            xlHoja1.Range("E" & Trim(Str(nfil)) & ":E" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp3(3))
            xlHoja1.Range("F" & Trim(Str(nfil)) & ":F" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp3(4))
            xlHoja1.Range("G" & Trim(Str(nfil)) & ":G" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp3(5))
            xlHoja1.Range("H" & Trim(Str(nfil)) & ":H" & Trim(Str(nfil))).Formula = "=" & ReduceFormulaSumatoria(nFilTemp3(6))
            
            'Cabecera
            ExcelCuadro xlHoja1, 1, nfil - 2, 8, nfil - 1
            'Cuadro del centro
            ExcelCuadro xlHoja1, 1, nfil, 8, nfil - 1
            'Cuadro del total
            ExcelCuadro xlHoja1, 2, nfil, 8, nfil
            For i = 1 To 6
                nFilTemp3(i) = ""
            Next
           
           xlHoja1.Cells.Font.Name = "Arial"
           xlHoja1.Cells.Font.Size = 8
           xlHoja1.Cells.EntireColumn.AutoFit
            
           ExcelEnd sArchGrabar, xlAplicacion, xlLibro, xlHoja1
           lbLibroOpen = False
            
            Screen.MousePointer = vbDefault
            MsgBox "Reporte Generado Satisfactoriamente", vbInformation, "Aviso!!!"
            CargaArchivo "PagaresxVenc" & Format(txtfechaDel.Text, "ddMMYYYY") & "_" & Format(txtFechaAl.Text, "ddMMYYYY") & ".XLSx", App.path & "\SPOOLER"
            
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Grabo la Operaci?n "
                Set objPista = Nothing
                '****
        End If
    Else
        Screen.MousePointer = vbDefault
        MsgBox "No hay datos para mostrar reporte", vbInformation, "Aviso!!!"
    End If
    rs.Close
    Set rs = Nothing
     
    
Exit Sub
ErrBegin:
  
  'ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
  Screen.MousePointer = vbDefault
  MsgBox TextErr(Err.Description), vbInformation, "Aviso"
     
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
    CierraConexion
End Sub

Private Sub txtCodObjeto_EmiteDatos()
If txtCodObjeto <> "" Then
    lblObjDesc = txtCodObjeto.psDescripcion
    cmdGenerar.SetFocus
End If
End Sub

Private Sub txtFechaAl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.chktodos.SetFocus
    End If
End Sub

Private Sub txtFechaDel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFechaAl.SetFocus
    End If
End Sub

Public Function CargaDatosPagaresxFecha(lsPersCod As String, lsFiltroCta As String, Optional lsPersCodFin As String = "", Optional lnIndiceVac As Double, Optional pdfecha1 As Date, Optional pdFecha2 As Date, Optional psOpeCod As String = "") As ADODB.Recordset
Dim lsFiltro As String
Dim sSql As String
Dim oCon As DConecta

Set oCon = New DConecta

On Error GoTo CargaDatosGeneralesCtaIFErr
    
    If lsPersCod <> "" Then
        lsFiltro = " AND ci.cPersCod = '" & Mid(lsPersCod, 4, 13) & "' "
    End If
    If lsPersCodFin <> "" Then
        If Len(lsPersCodFin) > 16 Then
            lsFiltro = " AND ci.cPersCod BETWEEN '" & Mid(lsPersCod, 4, 13) & "' and '" & Mid(lsPersCodFin, 4, 13) & "' and ci.cCtaIFCod BETWEEN '" & Mid(lsPersCod, 18, 10) & "' and '" & Mid(lsPersCodFin, 18, 10) & "'"
        Else
            lsFiltro = " AND ci.cPersCod BETWEEN '" & Mid(lsPersCod, 4, 13) & "' and '" & Mid(lsPersCodFin, 4, 13) & "' "
        End If
    End If
    
   sSql = " "
'   sSql = "SELECT ci.cIFTpo, ci.cPersCod, ci.cCtaIFCod, P.cPersNombre, ci.cCtaIFDesc, ci.dCtaIFAper, dCtaIFVenc, cia.nMontoPrestado, ci.nCtaIFPlazo, cia.nCtaIFCuotas, cia.nPeriodoGracia, cii.nCtaIFIntPeriodo, cii.nCtaIFIntValor nTasaInteres, cic.nNroCuota, cic.nInteresPagado + " _
'        & "             ISNULL( (SELECT SUM(" & IIf(Mid(psOpeCod, 3, 1) = "2", "me.nMovMEImporte", "mc.nMovImporte") & ") " _
'        & "              FROM Mov m JOIN MovCta mc on mc.nMovNro = m.nMovNro " & IIf(Mid(psOpeCod, 3, 1) = "2", "JOIN MovME me on me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem ", "") _
'        & "                       JOIN MovObjIF mif ON mif.nMovNro = mc.nMovNro and mif.nMovItem = mc.nMovItem JOIN OpeCta oc ON mc.cCtaContCod LIKE oc.cCtaContCod + '%' " _
'        & "              WHERE m.nMovFlag = 0 and m.cmovnro > '" & Format(pdFecha + 1, gsFormatoMovFecha) & "' and oc.copecod = '" & psOpeCod & "' and oc.cOpeCtaOrden = '3' and mif.cPersCod = ci.cPersCod and mif.cCtaIFCod = ci.cCtaIFCod and mif.cIFTpo = ci.cIFTpo " _
'        & "             ), 0) nInteresPagado, cic.dVencimiento, cia.nSaldoCap nSaldoCapCal, ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") nVac, cic.cMovNro, cia.cMonedaPago, " _
'        & "       Round( cia.nSaldoCap * CASE WHEN SubString(ci.cCtaIFCod,3,1) = '1' and cia.cMonedaPago = '2' THEN ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") ELSE 1 END,2) + " _
'        & "             ISNULL( (SELECT SUM(" & IIf(Mid(psOpeCod, 3, 1) = "2", "me.nMovMEImporte", "mc.nMovImporte") & ") " _
'        & "              FROM Mov m JOIN MovCta mc on mc.nMovNro = m.nMovNro " & IIf(Mid(psOpeCod, 3, 1) = "2", "JOIN MovME me on me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem ", "") _
'        & "                       JOIN MovObjIF mif ON mif.nMovNro = mc.nMovNro and mif.nMovItem = mc.nMovItem JOIN OpeCta oc ON mc.cCtaContCod LIKE oc.cCtaContCod + '%' " _
'        & "              WHERE m.nMovFlag = 0 and m.cmovnro > '" & Format(pdFecha + 1, gsFormatoMovFecha) & "' and oc.copecod = '" & psOpeCod & "' and oc.cOpeCtaOrden = '0' and mif.cPersCod = ci.cPersCod and mif.cCtaIFCod = ci.cCtaIFCod and mif.cIFTpo = ci.cIFTpo " _
'        & "             ), 0) nSaldoCap , ISNULL(cia.cCodLinCred,'') cCodLinCred, ISNULL(l.cDescripcion,'') cDesLinCred, cia.nCuotaPagoCap " _
'        & "    FROM CtaIF ci LEFT JOIN CtaIfAdeudados cia ON cia.cIFTpo = ci.cIFTpo and cia.cPersCod = ci.cPersCod and cia.cCtaIFCod = ci.cCtaIFCod LEFT JOIN ColocLineaCredito l ON l.cLineaCred = cia.cCodLinCred JOIN Persona P ON ci.cPersCod = p.cPersCod " _
'        & "         LEFT JOIN IndiceVac iv ON iv.dIndiceVac = ISNULL(cia.dCuotaUltPago, ci.dCtaIFAper) JOIN CtaIFInteres cii ON cii.cIFTpo = ci.cIFTpo and cii.cPersCod = ci.cPersCod and cii.cCtaIFCod = ci.cCtaIFCod " _
'        & "              and cii.dCtaIFIntRegistro = (SELECT Max(dCtaIFIntRegistro) " _
'        & "                                     FROM CtaIFInteres cii1 WHERE cii1.cIFTpo = cii.cIFTpo " _
'        & "                                      and cii1.cPersCod = cii.cPersCod and cii1.cCtaIFCod = cii.cCtaIFCod ) " _
'        & "         LEFT JOIN CtaIFCalendario cic ON cic.cIFTpo = ci.cIFTpo and cic.cPersCod = ci.cPersCod and cic.cCtaIFCod = ci.cCtaIFCod " _
'        & "              and cic.cTpoCuota = '2' and cic.nNroCuota = (SELECT Min(nNroCuota) FROM CtaIFCalendario cic1 " _
'        & "                        Where cic1.cIFTpo = cic.cIFTpo And cic1.cPersCod = cic.cPersCod And cic1.cCtaIFCod = cic.cCtaIFCod " _
'        & "                          and cic1.cTpoCuota = cic.cTpoCuota and cEstado = 0 ) " _
'        & "    WHERE ci.cCtaIFEstado in (" & gEstadoCtaIFActiva & "," & gEstadoCtaIFRegistrada & ") and  ci.cIFTpo+ci.cCtaIFCod LIKE '" & lsFiltroCta & "' and datediff(d,ci.dCtaIFAper,'" & Format(pdFecha, gsFormatoFecha) & "') >= 0 " & lsFiltro _
'        & "    ORDER BY ci.cIFtpo, ci.cPersCod, ISNULL(cia.cCodLinCred,''), ci.cCtaIFDesc  "
    '''''''''''''''''''''''
'    sSql = "SELECT convert(varchar(8), cic.dVencimiento,112) as dFechaVenc, ci.cIFTpo, ci.cPersCod, ci.cCtaIFCod, P.cPersNombre, ci.cCtaIFDesc, cii.nCtaIFIntValor nTasaInteres, " & _
'         " isnull(round( cic.nInteres * CASE WHEN SubString(ci.cCtaIFCod,3,1) = '1' and cia.cMonedaPago = '2' THEN ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") ELSE 1 END,2),0) as nInteresCon_VAC, " & _
'         " isnull(round( cic.nCapital * CASE WHEN SubString(ci.cCtaIFCod,3,1) = '1' and cia.cMonedaPago = '2' THEN ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") ELSE 1 END,2),0) as nCapitalCon_VAC, " & _
'         " isnull(round( cic.nComision * CASE WHEN SubString(ci.cCtaIFCod,3,1) = '1' and cia.cMonedaPago = '2' THEN ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") ELSE 1 END,2),0) as nComisionCon_VAC, " & _
'         " ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") nVac, cic.cMovNro, cia.cMonedaPago, " & _
'         " Round( cia.nSaldoCap * CASE WHEN SubString(ci.cCtaIFCod,3,1) = '1' and cia.cMonedaPago = '2' " & _
'         " THEN ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") ELSE 1 END,2) + ISNULL( (SELECT SUM(mc.nMovImporte) " & _
'         " FROM Mov m JOIN MovCta mc on mc.nMovNro = m.nMovNro JOIN MovObjIF mif ON mif.nMovNro = mc.nMovNro and mif.nMovItem = mc.nMovItem " & _
'         " JOIN OpeCta oc ON mc.cCtaContCod LIKE oc.cCtaContCod + '%'  " & _
'         " WHERE m.nMovFlag = 0 and m.cmovnro > '" & Format(pdFecha2 + 1, gsFormatoMovFecha) & "' and oc.copecod = '" & psOpeCod & "' " & _
'         " and oc.cOpeCtaOrden = '0' and mif.cPersCod = ci.cPersCod and mif.cCtaIFCod = ci.cCtaIFCod and mif.cIFTpo = ci.cIFTpo), 0) nSaldoCap, " & _
'         " cic.nNroCuota, cic.dVencimiento, ISNULL(cia.cCodLinCred,'') cCodLinCred, ISNULL(l.cDescripcion,'') cDesLinCred, cia.nCuotaPagoCap " & _
'         " FROM CtaIF ci LEFT JOIN CtaIfAdeudados cia ON cia.cIFTpo = ci.cIFTpo and cia.cPersCod = ci.cPersCod and cia.cCtaIFCod = ci.cCtaIFCod " & _
'         " LEFT JOIN ColocLineaCredito l ON l.cLineaCred = cia.cCodLinCred JOIN Persona P ON ci.cPersCod = p.cPersCod " & _
'         " LEFT JOIN IndiceVac iv ON iv.dIndiceVac = ISNULL(cia.dCuotaUltPago, " & _
'         " ci.dCtaIFAper) JOIN CtaIFInteres cii ON cii.cIFTpo = ci.cIFTpo and cii.cPersCod = ci.cPersCod and cii.cCtaIFCod = ci.cCtaIFCod " & _
'         " and cii.dCtaIFIntRegistro = (SELECT Max(dCtaIFIntRegistro) FROM CtaIFInteres cii1 WHERE cii1.cIFTpo = cii.cIFTpo " & _
'         " and cii1.cPersCod = cii.cPersCod and cii1.cCtaIFCod = cii.cCtaIFCod ) " & _
'         " LEFT JOIN CtaIFCalendario cic ON cic.cIFTpo = ci.cIFTpo and cic.cPersCod = ci.cPersCod and cic.cCtaIFCod = ci.cCtaIFCod " & _
'         " and cic.cTpoCuota = '2' and cic.nNroCuota >= (SELECT Min(nNroCuota) FROM CtaIFCalendario cic1 " & _
'         " Where cic1.cIFTpo = cic.cIFTpo And cic1.cPersCod = cic.cPersCod And cic1.cCtaIFCod = cic.cCtaIFCod and cic1.cTpoCuota = cic.cTpoCuota and cEstado = 0 ) " & _
'         " WHERE ci.cCtaIFEstado in (" & gEstadoCtaIFActiva & "," & gEstadoCtaIFRegistrada & ") and  ci.cIFTpo+ci.cCtaIFCod LIKE '" & lsFiltroCta & "' " & _
'         " and (convert(varchar(8), cic.dVencimiento,112) Between '" & Format(pdFecha1, "YYYYMMdd") & "' and '" & Format(pdFecha2, "YYYYMMdd") & "') " & lsFiltro & _
'         " ORDER BY convert(varchar(8), cic.dVencimiento,112), ci.cIFtpo, ci.cPersCod, ISNULL(cia.cCodLinCred,''), ci.cCtaIFDesc "
      
       sSql = "SELECT convert(varchar(8), cic.dVencimiento,112) as dFechaVenc, ci.cIFTpo, ci.cPersCod, ci.cCtaIFCod, P.cPersNombre, ci.cCtaIFDesc, cii.nCtaIFIntValor nTasaInteres, " & _
         " isnull(round( cic.nInteres * CASE WHEN SubString(ci.cCtaIFCod,3,1) = '1' and cia.cMonedaPago = '2' THEN ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") ELSE 1 END,2),0) as nInteresCon_VAC, " & _
         " isnull(round( cic.nCapital * CASE WHEN SubString(ci.cCtaIFCod,3,1) = '1' and cia.cMonedaPago = '2' THEN ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") ELSE 1 END,2),0) as nCapitalCon_VAC, " & _
         " isnull(round( cic.nComision * CASE WHEN SubString(ci.cCtaIFCod,3,1) = '1' and cia.cMonedaPago = '2' THEN ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") ELSE 1 END,2),0) as nComisionCon_VAC, " & _
         " ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") nVac, cic.cMovNro, cia.cMonedaPago, " & _
         " Round( cia.nSaldoCap * CASE WHEN SubString(ci.cCtaIFCod,3,1) = '1' and cia.cMonedaPago = '2' " & _
         " THEN ISNULL(iv.nIndiceVac, " & lnIndiceVac & ") ELSE 1 END,2) + ISNULL( (SELECT SUM(mc.nMovImporte) " & _
         " FROM Mov m JOIN MovCta mc on mc.nMovNro = m.nMovNro JOIN MovObjIF mif ON mif.nMovNro = mc.nMovNro and mif.nMovItem = mc.nMovItem " & _
         " JOIN OpeCta oc ON mc.cCtaContCod LIKE oc.cCtaContCod + '%'  " & _
         " WHERE m.nMovFlag = 0 and m.cmovnro > '" & Format(pdFecha2 + 1, gsFormatoMovFecha) & "' and oc.copecod = '" & psOpeCod & "' " & _
         " and oc.cOpeCtaOrden = '0' and mif.cPersCod = ci.cPersCod and mif.cCtaIFCod = ci.cCtaIFCod and mif.cIFTpo = ci.cIFTpo), 0) nSaldoCap, " & _
         " cic.nNroCuota, cic.dVencimiento, ISNULL(cia.cCodLinCred,'') cCodLinCred, ISNULL(l.cDescripcion,'') cDesLinCred, cia.nCuotaPagoCap " & _
         " FROM CtaIF ci LEFT JOIN CtaIfAdeudados cia ON cia.cIFTpo = ci.cIFTpo and cia.cPersCod = ci.cPersCod and cia.cCtaIFCod = ci.cCtaIFCod " & _
         " LEFT JOIN ColocLineaCredito l ON l.cLineaCred = cia.cCodLinCred JOIN Persona P ON ci.cPersCod = p.cPersCod "
         '" LEFT JOIN IndiceVac iv ON iv.dIndiceVac = ISNULL(cia.dCuotaUltPago, ci.dCtaIFAper)
    sSql = sSql & " JOIN CtaIFInteres cii ON cii.cIFTpo = ci.cIFTpo and cii.cPersCod = ci.cPersCod and cii.cCtaIFCod = ci.cCtaIFCod " & _
         " and cii.dCtaIFIntRegistro = (SELECT Max(dCtaIFIntRegistro) FROM CtaIFInteres cii1 WHERE cii1.cIFTpo = cii.cIFTpo " & _
         " and cii1.cPersCod = cii.cPersCod and cii1.cCtaIFCod = cii.cCtaIFCod ) " & _
         " LEFT JOIN CtaIFCalendario cic ON cic.cIFTpo = ci.cIFTpo and cic.cPersCod = ci.cPersCod and cic.cCtaIFCod = ci.cCtaIFCod " & _
         " and cic.cTpoCuota = '2' and cic.nNroCuota >= (SELECT Min(nNroCuota) FROM CtaIFCalendario cic1 " & _
         " Where cic1.cIFTpo = cic.cIFTpo And cic1.cPersCod = cic.cPersCod And cic1.cCtaIFCod = cic.cCtaIFCod and cic1.cTpoCuota = cic.cTpoCuota and cEstado = 0 ) "
         
    sSql = sSql & " LEFT JOIN IndiceVac iv ON iv.dIndiceVac = cic.dVencimiento "
         
    sSql = sSql & " WHERE ci.cCtaIFEstado in (" & gEstadoCtaIFActiva & "," & gEstadoCtaIFRegistrada & ") and  ci.cIFTpo+ci.cCtaIFCod LIKE '" & lsFiltroCta & "' " & _
         " and (convert(varchar(8), cic.dVencimiento,112) Between '" & Format(pdfecha1, "YYYYMMdd") & "' and '" & Format(pdFecha2, "YYYYMMdd") & "') " & lsFiltro & _
         " ORDER BY convert(varchar(8), cic.dVencimiento,112), ci.cIFtpo, ci.cPersCod, ISNULL(cia.cCodLinCred,''), ci.cCtaIFDesc "
      
      
oCon.AbreConexion
    Set CargaDatosPagaresxFecha = oCon.CargaRecordSet(sSql)
    oCon.CierraConexion
Exit Function
CargaDatosGeneralesCtaIFErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ReduceFormulaSumatoria(psSumatoria As String) As String
    Dim lsCadena As String
    Dim lsCadenaSub As String
    Dim lnPosSuma As Integer
    Dim lnAnt As Currency
    Dim lsCarAnt As String
    Dim lnPosNum As Integer
    
    Dim lsRes As String
    
    lsCadena = psSumatoria
    
    lnAnt = 0
    While lsCadena <> ""
        lnPosSuma = InStr(1, lsCadena, "+", vbTextCompare)
        If lnPosSuma > 0 Then
            lsCadenaSub = Left(lsCadena, lnPosSuma - 1)
            lsCadena = Mid(lsCadena, lnPosSuma + 1)
            
            If lsCadenaSub <> "" Then
                If IsNumeric(Mid(lsCadenaSub, 2)) Then
                    lnPosNum = 2
                ElseIf IsNumeric(Mid(lsCadenaSub, 3)) Then
                    lnPosNum = 3
                End If
                
                If lnAnt = 0 Then
                    lsRes = "SUM(" & lsCadenaSub
                ElseIf lnAnt + 1 <> CInt(Mid(lsCadenaSub, lnPosNum)) Then
                    lsRes = lsRes & ":" & lsCarAnt & ")+SUM(" & lsCadenaSub
                End If
                
                lsCarAnt = lsCadenaSub
                lnAnt = CInt(Mid(lsCadenaSub, lnPosNum))
            End If
        Else
            If lsRes = "" Then
                lsRes = lsCadena
            Else
                lsRes = lsRes & ":" & lsCadena & ")"
            End If
            
            lsCadena = ""
        End If
    Wend
    ReduceFormulaSumatoria = lsRes
End Function

