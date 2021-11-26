VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmOperacionesNum 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Operaciones realizadas por Usuarios"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10755
   Icon            =   "frmOperacionesNum.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkRangoFecha 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Usar Rango"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5175
      TabIndex        =   19
      Top             =   52
      Width           =   1440
   End
   Begin VB.CheckBox chkDetalle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Mostrar Detalle"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7155
      TabIndex        =   18
      Top             =   45
      Width           =   1950
   End
   Begin VB.CheckBox chkPesos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Pesos Ponderados"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3135
      TabIndex        =   17
      Top             =   45
      Width           =   1950
   End
   Begin MSComctlLib.ProgressBar PBT 
      Height          =   210
      Left            =   45
      TabIndex        =   11
      Top             =   5955
      Visible         =   0   'False
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   370
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CheckBox chkTodosGrupos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Todos los grupos"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   105
      TabIndex        =   3
      Top             =   660
      Width           =   1875
   End
   Begin VB.ListBox lstGrupos 
      Appearance      =   0  'Flat
      Height          =   1380
      ItemData        =   "frmOperacionesNum.frx":030A
      Left            =   60
      List            =   "frmOperacionesNum.frx":030C
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   960
      Width           =   5325
   End
   Begin VB.CheckBox chkTodosAge 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "&Todas las Agencias"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5415
      TabIndex        =   5
      Top             =   660
      Width           =   1950
   End
   Begin VB.CommandButton cmdGenerarExcell 
      Caption         =   "Pasar a Excel   >>>"
      Enabled         =   0   'False
      Height          =   345
      Left            =   7620
      TabIndex        =   8
      Top             =   5580
      Width           =   1965
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      Height          =   345
      Left            =   6480
      TabIndex        =   7
      Top             =   5580
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   9660
      TabIndex        =   9
      Top             =   5580
      Width           =   1035
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
      Height          =   660
      Left            =   45
      TabIndex        =   10
      Top             =   -15
      Width           =   2760
      Begin MSMask.MaskEdBox mskIni 
         Height          =   300
         Left            =   105
         TabIndex        =   0
         Top             =   270
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFin 
         Height          =   300
         Left            =   1500
         TabIndex        =   2
         Top             =   270
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
   End
   Begin VB.ListBox lstAge 
      Appearance      =   0  'Flat
      Height          =   1380
      Left            =   5385
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   960
      Width           =   5325
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   3105
      Left            =   75
      TabIndex        =   1
      Top             =   2385
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   5477
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      TextStyleFixed  =   4
      MergeCells      =   1
      AllowUserResizing=   3
      Appearance      =   0
      RowSizingMode   =   1
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSComctlLib.ProgressBar PBA 
      Height          =   210
      Left            =   45
      TabIndex        =   12
      Top             =   6390
      Visible         =   0   'False
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   370
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   345
      Left            =   8520
      TabIndex        =   15
      Top             =   5580
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.OLE OleExcel 
      Appearance      =   0  'Flat
      AutoActivate    =   3  'Automatic
      Enabled         =   0   'False
      Height          =   255
      Left            =   8985
      SizeMode        =   1  'Stretch
      TabIndex        =   16
      Top             =   435
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label lblAge 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "lblAge"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3000
      TabIndex        =   14
      Top             =   6195
      Visible         =   0   'False
      Width           =   5100
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4305
      TabIndex        =   13
      Top             =   5700
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "frmOperacionesNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsUsuario As String
Dim lnPosTotal As Integer
Dim lnCamposNum As Integer

Dim lnMes As Integer
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim xlHoja2 As Excel.Worksheet

Dim lnNumMes As Integer

Private Sub chkDirExcel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.chkTodosGrupos.SetFocus
    End If
End Sub

Private Sub chkTodosAge_Click()
    Dim i As Integer
    
    For i = 0 To Me.lstAge.ListCount - 1
        If chkTodosAge.value = 1 Then
            Me.lstAge.Selected(i) = True
        Else
            Me.lstAge.Selected(i) = False
        End If
    Next i
End Sub

Private Sub chkTodosAge_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.lstAge.SetFocus
    End If
End Sub

Private Sub chkTodosGrupos_Click()
    Dim i As Integer
    
    For i = 0 To Me.lstGrupos.ListCount - 1
        If chkTodosGrupos.value = 1 Then
            Me.lstGrupos.Selected(i) = True
        Else
            Me.lstGrupos.Selected(i) = False
        End If
    Next i
End Sub

Private Sub chkTodosGrupos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.lstGrupos.SetFocus
    End If
End Sub

Private Sub cmdGenerar_Click()
    Dim i As Integer
    Dim rsA As New ADODB.Recordset
    Dim rsD As New ADODB.Recordset
    
    If Not IsDate(Me.mskIni.Text) Then
        MsgBox "Debe ingresar una fecha de inicio valida.", vbInformation, "Aviso"
        Me.mskIni.SetFocus
        Exit Sub
    ElseIf Not IsDate(Me.mskFin.Text) Then
        MsgBox "Debe ingresar una fecha de inicio valida.", vbInformation, "Aviso"
        Me.mskFin.SetFocus
        Exit Sub
    End If
    
    SetFormato
        
    If Flex.TextMatrix(2, 0) = "" Or Flex.TextMatrix(0, 2) = "" Then Exit Sub
        
    ActivaBarras True
    
    Me.PBT.Min = 0
    Me.PBT.Max = Me.Flex.Cols - 2
    Me.PBT.value = Me.PBT.Min
    
 
    For i = 2 To Me.Flex.Cols - 2
    
        If Me.Flex.TextMatrix(0, i) <> "" And Me.Flex.TextMatrix(0, i) <> "TOTAL" Then
            DoEvents
            GetData Me.Flex.TextMatrix(0, i), i
            Me.PBT.value = Me.PBT.value + 1
        End If
    Next i
    
    Me.PBT.value = Me.PBT.Max
    
    MsgBox "EL Reporte a Finalizado.", vbInformation, "Aviso"
    
    ActivaBarras False
    
    CalculaSuma
    Me.cmdGenerarExcell.Enabled = True
End Sub

Private Sub cmdGenerarExcell_Click()
    Dim lsArchivoN  As String
    Dim lbLibroOpen As Boolean
    
    If Me.chkPesos.value = 0 Then
        lsArchivoN = App.path & "\Spooler\" & Format(CDate(Me.mskFin.Text), "yyyymmdd") & "OPENUM.xls"
    Else
        lsArchivoN = App.path & "\Spooler\" & Format(CDate(Me.mskFin.Text), "yyyymmdd") & "OPEPESOS.xls"
    End If
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
       If lbLibroOpen Then
          Set xlHoja1 = xlLibro.Worksheets(1)
          If Me.chkPesos.value = 0 Then
            ExcelAddHoja Format(gdFecSis, "yyyymmdd") & "_OPENUM", xlLibro, xlHoja1
          Else
            ExcelAddHoja Format(gdFecSis, "yyyymmdd") & "_OPEPESOS", xlLibro, xlHoja1
          End If
          Call GeneraReporteHoja1
          OleExcel.Class = "ExcelWorkSheet"
          ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1

          OleExcel.SourceDoc = lsArchivoN
          OleExcel.Verb = 1
          OleExcel.Action = 1
          OleExcel.DoVerb -1
       End If
End Sub



Private Sub cmdsalir_Click()
    'CalculaSuma
    Unload Me
    
End Sub

Private Sub Form_Load()
    Dim Sql As String
    Dim oAcceso As COMDPersona.UCOMAcceso
    Set oAcceso = New COMDPersona.UCOMAcceso
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sCadTemp As String
    Dim oGen As COMDConstSistema.DCOMGeneral
    Set oGen = New COMDConstSistema.DCOMGeneral
    
    
    Call oAcceso.CargaControlGrupos(gsDominio)
    sCadTemp = oAcceso.DameGrupo
    Do While sCadTemp <> ""
        lstGrupos.AddItem sCadTemp
        sCadTemp = oAcceso.DameGrupo
    Loop
    
    Set rs = oGen.GetNombreAgencias
    While Not rs.EOF
        lstAge.AddItem rs.Fields(0) & "  " & rs.Fields(1)
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    ActivaBarras False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'CierraConexion
End Sub

Private Sub SetFormato()
    Dim i As Integer
    Dim Sql As String
    Dim sqlAdd As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim lbBan As Boolean
    Dim k As Integer
    Dim oConSis As COMDConstSistema.DCOMOperacion
    Set oConSis = New COMDConstSistema.DCOMOperacion
    Dim lsTempUsu As String
    Dim lsTempUsuGrupo As String
    Dim oAccesso As COMDPersona.UCOMAcceso
    Set oAccesso = New COMDPersona.UCOMAcceso
    

    
    lnNumMes = DateDiff("m", CDate(Me.mskIni.Text), CDate(Me.mskFin.Text))
    
    If Me.chkRangoFecha.value = 1 Then
        lnNumMes = 0
    End If
    
    
    Flex.MergeCells = flexMergeRestrictAll
    
    'Me.Flex.Rows = 0
    'Me.Flex.Cols = 0
    
    Me.Flex.Rows = 4
    Me.Flex.FixedRows = 3
    Me.Flex.Cols = 3
    Me.Flex.FixedCols = 2
    
    Me.Flex.TextMatrix(0, 0) = "Usuario"
    Me.Flex.TextMatrix(0, 1) = "Nombre/Agencia "
    Me.Flex.TextMatrix(1, 0) = "Usuario"
    Me.Flex.TextMatrix(1, 1) = "Nombre/Agencia "
    Me.Flex.TextMatrix(2, 0) = "Usuario"
    Me.Flex.TextMatrix(2, 1) = "Nombre/Agencia "
    Me.Flex.ColWidth(0) = 800
    Me.Flex.ColWidth(1) = 3000
    
    Me.Flex.MergeRow(0) = True
    Me.Flex.MergeRow(1) = True
    
    Me.Flex.MergeCol(0) = True
    Me.Flex.MergeCol(1) = True
    Me.Flex.MergeCol(2) = True

    Me.Flex.ColAlignmentFixed(0) = 4
    Me.Flex.ColAlignmentFixed(1) = 1

    Set rs = oConSis.ObtenerValoresPonderados()
    
    
    lnCamposNum = rs.RecordCount
    
    For i = 0 To Me.lstAge.ListCount - 1
        rs.MoveFirst
        lbBan = False
        If Me.lstAge.Selected(i) Then
            For k = 0 To lnNumMes
                lbBan = False
                rs.MoveFirst
                
                If Me.Flex.TextMatrix(0, 2) = "" Then
                    Me.Flex.MergeCol(Me.Flex.Cols - 1) = True
                    Me.Flex.TextMatrix(0, Me.Flex.Cols - 1) = Trim(lstAge.List(i)) ' Left(lstAge.List(i), 5)
                    While Not rs.EOF
                        If lbBan Then
                            Me.Flex.Cols = Me.Flex.Cols + 1
                        Else
                            lbBan = True
                        End If
                        Me.Flex.Col = Me.Flex.Cols - 1
                        Me.Flex.MergeCol(Me.Flex.Cols - 1) = True
                        Me.Flex.Row = 0
                        Me.Flex.CellAlignment = 4
                        Me.Flex.Row = 1
                        Me.Flex.CellAlignment = 4
                        Me.Flex.Row = 2
                        Me.Flex.CellAlignment = 0
                        Me.Flex.TextMatrix(0, Me.Flex.Cols - 1) = Trim(lstAge.List(i)) ' Left(lstAge.List(i), 5)
                        
                        If Me.chkRangoFecha.value = 1 Then
                            Me.Flex.TextMatrix(1, Me.Flex.Cols - 1) = Me.mskIni.Text & " - " & Me.mskFin.Text & " - [" & Trim(Str(k)) & "]"
                        Else
                            Me.Flex.TextMatrix(1, Me.Flex.Cols - 1) = UCase(Format(DateAdd("M", k, CDate(Me.mskIni.Text)), "MMMM YYYY")) & " - [" & Trim(Str(k)) & "]"
                        End If
                        
                        Me.Flex.TextMatrix(2, Me.Flex.Cols - 1) = rs!cNomTab & Space(50) & Trim(rs!cValor)
                        Me.Flex.ColWidth(Me.Flex.Cols - 1) = 1500
                        rs.MoveNext
                    Wend
                Else
                    Me.Flex.Cols = Me.Flex.Cols + 1
                    Me.Flex.TextMatrix(0, Me.Flex.Cols - 1) = Trim(lstAge.List(i)) ' Left(lstAge.List(i), 5)
                    rs.MoveFirst
                    While Not rs.EOF
                        If lbBan Then
                            Me.Flex.Cols = Me.Flex.Cols + 1
                        Else
                            lbBan = True
                        End If
                        Me.Flex.Col = Me.Flex.Cols - 1
                        Me.Flex.MergeCol(Me.Flex.Cols - 1) = True
                        Me.Flex.Row = 0
                        Me.Flex.CellAlignment = 4
                        Me.Flex.Row = 1
                        Me.Flex.CellAlignment = 4
                        Me.Flex.Row = 2
                        Me.Flex.CellAlignment = 0
                        Me.Flex.TextMatrix(0, Me.Flex.Cols - 1) = Trim(lstAge.List(i)) ' Left(lstAge.List(i), 5)
                        Me.Flex.TextMatrix(1, Me.Flex.Cols - 1) = UCase(Format(DateAdd("M", k, CDate(Me.mskIni.Text)), "MMMM YYYY")) & " - [" & Trim(Str(k)) & "]"
                        Me.Flex.TextMatrix(2, Me.Flex.Cols - 1) = rs!cNomTab & Space(50) & Trim(rs!cValor)
                        Me.Flex.ColWidth(Me.Flex.Cols - 1) = 1500
                        rs.MoveNext
                    Wend
                End If
            Next k
        End If
    Next i
    
    If Me.Flex.TextMatrix(0, 3) = "" Then
        MsgBox "No ha elegido ninguna Agencia.", vbInformation, "Aviso"
        Me.lstAge.SetFocus
        Exit Sub
    End If
    
    Me.Flex.Cols = Me.Flex.Cols + 1
    lnPosTotal = Me.Flex.Cols
    Me.Flex.TextMatrix(0, Me.Flex.Cols - 1) = "TOTAL"
    rs.MoveFirst
    lbBan = False
    While Not rs.EOF
        If lbBan Then
            Me.Flex.Cols = Me.Flex.Cols + 1
        Else
            lbBan = True
        End If
        Me.Flex.Col = Me.Flex.Cols - 1
        Me.Flex.MergeCol(Me.Flex.Cols - 1) = True
        Me.Flex.MergeCol(Me.Flex.Cols - 1) = True
        Me.Flex.Row = 0
        Me.Flex.CellAlignment = 4
        Me.Flex.Row = 1
        Me.Flex.CellAlignment = 0
        Me.Flex.TextMatrix(0, Me.Flex.Cols - 1) = "TOTAL"
        Me.Flex.TextMatrix(1, Me.Flex.Cols - 1) = rs!cNomTab & Space(50) & Trim(rs!cValor)
        Me.Flex.TextMatrix(2, Me.Flex.Cols - 1) = rs!cNomTab & Space(50) & Trim(rs!cValor)
        Me.Flex.ColWidth(Me.Flex.Cols - 1) = 1500
        rs.MoveNext
    Wend
    
    rs.Close
    sqlAdd = "''"
    
    For i = 0 To Me.lstGrupos.ListCount - 1
        If Me.lstGrupos.Selected(i) Then
            If sqlAdd = "" Then
                sqlAdd = "'" & Trim(lstGrupos.List(i)) & "'"
            Else
                sqlAdd = sqlAdd & ",'" & Trim(lstGrupos.List(i)) & "'"
            End If
        End If
    Next i
    
    
    Call oAccesso.CargaControlUsuarios(gsDominio)
    lsTempUsu = oAccesso.DameUsuario
    lsUsuario = "''"
    
    While lsTempUsu <> ""
        Call oAccesso.CargaGruposUsuario(lsTempUsu, gsDominio)
        lsTempUsuGrupo = oAccesso.DameGrupoUsuario
        
        lbBan = False
        While lsTempUsuGrupo <> "" And Not lbBan
            If InStr(1, sqlAdd, lsTempUsuGrupo) Then lbBan = True
            lsTempUsuGrupo = oAccesso.DameGrupoUsuario
        Wend
        
        If lbBan Then
            If Me.Flex.TextMatrix(3, 0) = "" Then
                Me.Flex.TextMatrix(Me.Flex.Rows - 1, 0) = lsTempUsu
                Me.Flex.TextMatrix(Me.Flex.Rows - 1, 1) = oAccesso.MostarNombre(gsDominio, lsTempUsu)
                lsUsuario = "'" & lsTempUsu & "'"
            Else
                Me.Flex.Rows = Me.Flex.Rows + 1
                Me.Flex.TextMatrix(Me.Flex.Rows - 1, 0) = lsTempUsu
                Me.Flex.TextMatrix(Me.Flex.Rows - 1, 1) = oAccesso.MostarNombre(gsDominio, lsTempUsu)
                lsUsuario = lsUsuario & ",'" & lsTempUsu & "'"
            End If
        End If
        
        lsTempUsu = oAccesso.DameUsuario
    Wend
    
    If Me.Flex.TextMatrix(1, 0) = "" Then
        MsgBox "No ha elegido ninguna Agencia.", vbInformation, "Aviso"
        Me.lstGrupos.SetFocus
        Exit Sub
    End If
    
    Me.Flex.Rows = Me.Flex.Rows + 1
    Me.Flex.TextMatrix(Me.Flex.Rows - 1, 1) = "TOTAL"
    
End Sub

Private Sub ActivaBarras(pbActiva As Boolean)
    Me.lblTotal.Visible = pbActiva
    Me.lblAge.Visible = pbActiva
    Me.PBT.Visible = pbActiva
    Me.PBA.Visible = pbActiva
    
    If pbActiva Then
        Me.Height = 7005
    Else
        Me.Height = 6315
    End If
End Sub

Private Sub GetData(ByVal psAge As String, ByVal pnIndice As Integer)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim Sql As String
    Dim lsPeso As String
    Dim oConSis As COMDConstSistema.DCOMOperacion

    If Right(Me.Flex.TextMatrix(1, pnIndice), 4) = "9999" Then
        GetReporteTot psAge, pnIndice
    Else

            Me.PBA.Min = 0
            Me.PBA.value = Me.PBA.Min
            Me.lblAge.Caption = psAge

            DoEvents

            lnMes = CCur(Mid(Me.Flex.TextMatrix(1, pnIndice), InStr(1, Me.Flex.TextMatrix(1, pnIndice), "[", vbTextCompare) + 1, InStr(1, Me.Flex.TextMatrix(1, pnIndice), "]", vbTextCompare) - InStr(1, Me.Flex.TextMatrix(1, pnIndice), "[", vbTextCompare) - 1))

            If Me.chkRangoFecha.value = 1 Then
                If Me.chkPesos.value = 0 Then
                   Set oConSis = New COMDConstSistema.DCOMOperacion
                    Set rs = oConSis.ObtenerValoresPonderadosRelRangoFec(Me.mskIni.Text, Me.mskFin.Text, Trim(Me.Flex.TextMatrix(2, pnIndice)), lsUsuario)
                   Set oConSis = Nothing
                Else
                    lsPeso = Mid(Me.Flex.TextMatrix(1, pnIndice), InStr(1, Me.Flex.TextMatrix(1, pnIndice), "[") + 1, InStr(1, Me.Flex.TextMatrix(1, pnIndice), "]") - InStr(1, Me.Flex.TextMatrix(1, pnIndice), "[") - 1)
                    Set oConSis = New COMDConstSistema.DCOMOperacion
                        Set rs = oConSis.ObtenerValoresPonderadosRelRangoFec(Me.mskIni.Text, Me.mskFin.Text, Trim(Me.Flex.TextMatrix(2, pnIndice)), lsUsuario, lsPeso)
                    Set oConSis = Nothing

                End If
            Else
                If Me.chkPesos.value = 0 Then
                    Set oConSis = New COMDConstSistema.DCOMOperacion
                        Set rs = oConSis.ObtenerValoresPonderadosRelFec(Me.mskIni.Text, lnMes, Trim(Me.Flex.TextMatrix(2, pnIndice)), lsUsuario, "")
                    Set oConSis = Nothing

                Else
                    lsPeso = Mid(Me.Flex.TextMatrix(1, pnIndice), InStr(1, Me.Flex.TextMatrix(1, pnIndice), "[") + 1, InStr(1, Me.Flex.TextMatrix(1, pnIndice), "]") - InStr(1, Me.Flex.TextMatrix(1, pnIndice), "[") - 1)
                      Set oConSis = New COMDConstSistema.DCOMOperacion
                        Set rs = oConSis.ObtenerValoresPonderadosRelFec(Me.mskIni.Text, lnMes, Trim(Me.Flex.TextMatrix(2, pnIndice)), lsUsuario, lsPeso)
                    Set oConSis = Nothing

                End If
            End If
            If rs Is Nothing Then
            Else
                Me.PBA.Max = IIf(rs.RecordCount = 0, 1, rs.RecordCount)
                GetReporte rs, pnIndice
                Me.PBA.value = Me.PBA.Max
            End If

    End If
End Sub


Private Sub GetReporte(prs As ADODB.Recordset, pnIndice As Integer)
    Dim i As Integer
    Dim lnPos As Integer
    
    While Not prs.EOF
        For i = 1 To Me.Flex.Rows - 1
            If prs.Fields(0) = Me.Flex.TextMatrix(i, 0) Then
                lnPos = i
                i = Me.Flex.Rows - 1
            End If
        Next i
        
        Me.Flex.TextMatrix(lnPos, pnIndice) = Format(prs.Fields(1), "#,##0")
        
        prs.MoveNext
        Me.PBA.value = Me.PBA.value + 1
        DoEvents
    Wend

End Sub

Private Sub GetReporteTot(psAge As String, pnIndice As Integer)
    Dim i As Integer
    Dim J As Integer
    Dim lnPos As Integer
    Dim lsAcumulador As Currency
    
    For i = 2 To Me.Flex.Rows - 1
        lsAcumulador = 0
        For J = 2 To pnIndice
            If Flex.TextMatrix(0, J) = psAge And Flex.TextMatrix(i, J) <> "" Then
                lsAcumulador = lsAcumulador + CCur(Flex.TextMatrix(i, J))
            End If
        Next J
        Me.Flex.TextMatrix(i, pnIndice) = Format(lsAcumulador, "#,##0")
    Next i
    
    DoEvents
End Sub


Private Sub lstAge_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdGenerar.SetFocus
    End If
End Sub

Private Sub lstGrupos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.chkTodosAge.SetFocus
    End If
End Sub

Private Sub mskIni_GotFocus()
    mskIni.SelStart = 0
    mskIni.SelLength = 50
End Sub

Private Sub mskIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mskFin.SetFocus
    End If
End Sub

Private Sub mskFin_GotFocus()
    mskFin.SelStart = 0
    mskFin.SelLength = 50
End Sub

Private Sub mskFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.chkTodosGrupos.SetFocus
    End If
End Sub

Private Sub CalculaSuma()
    Dim i As Integer
    Dim J As Integer
    Dim k As Integer
    Dim lnContenedor As Currency
    
        For i = 3 To Me.Flex.Rows - 1
            For k = lnPosTotal - 1 To Me.Flex.Cols - 1
                lnContenedor = 0
                For J = 2 To lnPosTotal
                    If Me.Flex.TextMatrix(2, J) = Me.Flex.TextMatrix(2, k) Then
                        If Me.Flex.TextMatrix(i, J) <> "" Then
                            lnContenedor = lnContenedor + Me.Flex.TextMatrix(i, J)
                        End If
                    End If
                Next J
                Flex.TextMatrix(i, k) = Format(lnContenedor, "#,##0")
            Next k
        Next i
    

    For i = 2 To Me.Flex.Cols - 1
        lnContenedor = 0
        For J = 3 To Me.Flex.Rows - 2
            If Me.Flex.TextMatrix(J, i) <> "" Then
                lnContenedor = lnContenedor + Me.Flex.TextMatrix(J, i)
            End If
        Next J
        Flex.TextMatrix(Me.Flex.Rows - 1, i) = Format(lnContenedor, "#,##0")
    Next i
    
End Sub

'********************************
' Adiciona Hoja a LibroExcel
'********************************
Private Sub ExcelAddHoja(psHojName As String, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet)
For Each xlHoja1 In xlLibro.Worksheets
    If xlHoja1.Name = psHojName Then
       xlHoja1.Delete
       Exit For
    End If
Next
Set xlHoja1 = xlLibro.Worksheets.Add
xlHoja1.Name = psHojName
End Sub

Private Sub GeneraReporteHoja1()
    Dim i As Integer
    Dim k As Integer
    Dim J As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lnAcum As Currency
    Dim lnPosI As Integer
    Dim lnPosJ As Integer
    
    Dim sTipoGara As String
    Dim sTipoCred As String
    
    Dim lbUsu As Boolean
    Dim LbNom As Boolean
    Dim lbTot As Boolean
    
    Dim lsCadAux As String
    Dim lsCadAuxMes As String
   
    lnPosI = 0
    lnPosJ = 0
    lbUsu = False
    LbNom = False
    lbTot = False
    
    If Me.chkDetalle.value = 0 Then
        For i = 0 To Me.Flex.Rows - 1
            lnAcum = 0
            If Me.Flex.TextMatrix(i, Flex.Cols - 1) <> "0" Then
                lnPosI = lnPosI + 1
                lnPosJ = 0
                For J = 0 To Me.Flex.Cols - 1
                    If UCase(Me.Flex.TextMatrix(0, J)) = "TOTAL" Or UCase(Me.Flex.TextMatrix(0, J)) = "USUARIO" Or UCase(Me.Flex.TextMatrix(0, J)) = "NOMBRE/AGENCIA " Then
                        lnPosJ = lnPosJ + 1
                        If IsNumeric(Me.Flex.TextMatrix(i, J)) Then
                            xlHoja1.Cells(lnPosI, lnPosJ) = Format(Me.Flex.TextMatrix(i, J), "#,##0.00")
                        Else
                            If Not lbTot And (UCase(Me.Flex.TextMatrix(i, J)) = "TOTAL" Or UCase(Me.Flex.TextMatrix(i, J)) = "USUARIO" Or UCase(Me.Flex.TextMatrix(i, J)) = "NOMBRE/AGENCIA ") Then
                                xlHoja1.Cells(lnPosI, lnPosJ) = Trim(Mid(Me.Flex.TextMatrix(i, J), 1, 45))
                            ElseIf Not lbUsu And (UCase(Me.Flex.TextMatrix(i, J)) = "TOTAL" Or UCase(Me.Flex.TextMatrix(i, J)) = "USUARIO" Or UCase(Me.Flex.TextMatrix(i, J)) = "NOMBRE/AGENCIA ") Then
                                xlHoja1.Cells(lnPosI, lnPosJ) = Trim(Mid(Me.Flex.TextMatrix(i, J), 1, 45))
                            ElseIf Not LbNom And (UCase(Me.Flex.TextMatrix(i, J)) = "TOTAL" Or UCase(Me.Flex.TextMatrix(i, J)) = "USUARIO" Or UCase(Me.Flex.TextMatrix(i, J)) = "NOMBRE/AGENCIA ") Then
                                xlHoja1.Cells(lnPosI, lnPosJ) = Trim(Mid(Me.Flex.TextMatrix(i, J), 1, 45))
                            ElseIf UCase(Me.Flex.TextMatrix(i, J)) <> "TOTAL" And UCase(Me.Flex.TextMatrix(i, J)) <> "USUARIO" And UCase(Me.Flex.TextMatrix(i, J)) <> "NOMBRE/AGENCIA " Then
                                xlHoja1.Cells(lnPosI, lnPosJ) = Trim(Mid(Me.Flex.TextMatrix(i, J), 1, 45))
                            End If
                        End If
                        
                        If UCase(Me.Flex.TextMatrix(0, J)) = "TOTAL" Then
                            lbTot = True
                        ElseIf UCase(Me.Flex.TextMatrix(0, J)) = "USUARIO" Then
                            lbUsu = True
                        ElseIf UCase(Me.Flex.TextMatrix(0, J)) = "NOMBRE/AGENCIA " Then
                            LbNom = True
                        End If
                        
                    End If
                Next J
            End If
        Next i
            
        xlHoja1.Range("A1:A" & Trim(Str(Me.Flex.Rows))).Font.Bold = True
        xlHoja1.Range("B1:B" & Trim(Str(Me.Flex.Rows))).Font.Bold = True
        
        'Mesclar celdas
        xlHoja1.Range("A1:A3").Merge
        xlHoja1.Range("B1:B3").Merge
        xlHoja1.Range("C1:I1").Merge
        
        'Lineas
        xlHoja1.Range("A1:A" & Trim(Str(lnPosI))).BorderAround 1, xlMedium
        xlHoja1.Range("B1:B" & Trim(Str(lnPosI))).BorderAround 1, xlMedium
        xlHoja1.Range("C1:C" & Trim(Str(lnPosI))).BorderAround 1, xlMedium
        xlHoja1.Range("D1:D" & Trim(Str(lnPosI))).BorderAround 1, xlMedium
        xlHoja1.Range("E1:E" & Trim(Str(lnPosI))).BorderAround 1, xlMedium
        xlHoja1.Range("F1:F" & Trim(Str(lnPosI))).BorderAround 1, xlMedium
        xlHoja1.Range("G1:G" & Trim(Str(lnPosI))).BorderAround 1, xlMedium
        xlHoja1.Range("H1:H" & Trim(Str(lnPosI))).BorderAround 1, xlMedium
        xlHoja1.Range("I1:I" & Trim(Str(lnPosI))).BorderAround 1, xlMedium
         
        xlHoja1.Range("A1:I" & Trim(Str(1))).BorderAround 1, xlMedium
        xlHoja1.Range("A2:I" & Trim(Str(2))).BorderAround 1, xlMedium
        xlHoja1.Range("A3:I" & Trim(Str(2))).BorderAround 1, xlMedium
        
        xlHoja1.Range("A" & Trim(Str(lnPosI)) & ":I" & Trim(Str(lnPosI))).BorderAround 1, xlMedium
        
        xlHoja1.Range("A1:I" & Trim(Str(lnPosI))).NumberFormat = "#,##0.00"
        
        xlHoja1.Cells.Select
        xlHoja1.Cells.EntireColumn.AutoFit
    
        xlHoja1.Range("A4:I" & Trim(Str(lnPosI - 1))).Select
        xlHoja1.Range("A4:I" & Trim(Str(lnPosI - 1))).Sort xlHoja1.Range("I3"), xlDescending
        
        xlHoja1.Range("C1:I1").HorizontalAlignment = xlCenter
        xlHoja1.Range("A1:A2").HorizontalAlignment = xlCenter
        xlHoja1.Range("B1:B2").HorizontalAlignment = xlCenter
        xlHoja1.Range("A1:A2").VerticalAlignment = xlCenter
        xlHoja1.Range("B1:B2").VerticalAlignment = xlCenter
    Else
        For i = 0 To Me.Flex.Rows - 1
            lnAcum = 0
            If Me.Flex.TextMatrix(i, Flex.Cols - 1) <> "0" Then
                lnPosI = lnPosI + 1
                lnPosJ = 0
                For J = 0 To Me.Flex.Cols - 1
                    'If UCase(Me.Flex.TextMatrix(0, J)) = "TOTAL" Or UCase(Me.Flex.TextMatrix(0, J)) = "USUARIO" Or UCase(Me.Flex.TextMatrix(0, J)) = "NOMBRE/AGENCIA " Then
                        lnPosJ = lnPosJ + 1
                        If IsNumeric(Me.Flex.TextMatrix(i, J)) Then
                            xlHoja1.Cells(lnPosI, lnPosJ) = Format(Me.Flex.TextMatrix(i, J), "#,##0.00")
                        Else
                            If lsCadAux = Trim(Mid(Me.Flex.TextMatrix(i, J), 1, 45)) And i = 0 Then
                                xlHoja1.Cells(lnPosI, lnPosJ) = ""
                            ElseIf lsCadAuxMes = Trim(Mid(Me.Flex.TextMatrix(i, J), 1, 45)) And i = 1 Then
                                xlHoja1.Cells(lnPosI, lnPosJ) = ""
                            Else
                                If Not lbTot And (UCase(Me.Flex.TextMatrix(i, J)) = "TOTAL" Or UCase(Me.Flex.TextMatrix(i, J)) = "USUARIO" Or UCase(Me.Flex.TextMatrix(i, J)) = "NOMBRE/AGENCIA ") Then
                                    xlHoja1.Cells(lnPosI, lnPosJ) = Trim(Mid(Me.Flex.TextMatrix(i, J), 1, 45))
                                ElseIf Not lbUsu And (UCase(Me.Flex.TextMatrix(i, J)) = "TOTAL" Or UCase(Me.Flex.TextMatrix(i, J)) = "USUARIO" Or UCase(Me.Flex.TextMatrix(i, J)) = "NOMBRE/AGENCIA ") Then
                                    xlHoja1.Cells(lnPosI, lnPosJ) = Trim(Mid(Me.Flex.TextMatrix(i, J), 1, 45))
                                ElseIf Not LbNom And (UCase(Me.Flex.TextMatrix(i, J)) = "TOTAL" Or UCase(Me.Flex.TextMatrix(i, J)) = "USUARIO" Or UCase(Me.Flex.TextMatrix(i, J)) = "NOMBRE/AGENCIA ") Then
                                    xlHoja1.Cells(lnPosI, lnPosJ) = Trim(Mid(Me.Flex.TextMatrix(i, J), 1, 45))
                                ElseIf UCase(Me.Flex.TextMatrix(i, J)) <> "TOTAL" And UCase(Me.Flex.TextMatrix(i, J)) <> "USUARIO" And UCase(Me.Flex.TextMatrix(i, J)) <> "NOMBRE/AGENCIA " Then
                                    xlHoja1.Cells(lnPosI, lnPosJ) = Trim(Mid(Me.Flex.TextMatrix(i, J), 1, 45))
                                End If
                            End If
                            
                            lsCadAux = Trim(Mid(Me.Flex.TextMatrix(i, J), 1, 45))
                            lsCadAuxMes = Trim(Mid(Me.Flex.TextMatrix(i, J), 1, 45))
                        End If
                        
                        If UCase(Me.Flex.TextMatrix(0, J)) = "TOTAL" Then
                            lbTot = True
                        ElseIf UCase(Me.Flex.TextMatrix(0, J)) = "USUARIO" Then
                            lbUsu = True
                        ElseIf UCase(Me.Flex.TextMatrix(0, J)) = "NOMBRE/AGENCIA " Then
                            LbNom = True
                        End If
                        
                    'End If
                Next J
            End If
        Next i
        
        xlHoja1.Range("A1:A" & Trim(Str(Me.Flex.Rows))).Font.Bold = True
        xlHoja1.Range("B1:B" & Trim(Str(Me.Flex.Rows))).Font.Bold = True
        
        xlHoja1.Range("A1:A3").Merge
        xlHoja1.Range("B1:B3").Merge
        
        xlHoja1.Range("A1:A3").HorizontalAlignment = xlCenter
        xlHoja1.Range("B1:B3").HorizontalAlignment = xlCenter
        xlHoja1.Range("A1:A3").VerticalAlignment = xlCenter
        xlHoja1.Range("B1:B3").VerticalAlignment = xlCenter
        
        
        For i = 3 To (lnPosJ - lnCamposNum) Step ((lnCamposNum) * (lnMes + 1))
            xlHoja1.Range(ExcelColumnaString(i) & "1:" & ExcelColumnaString(i + (lnCamposNum * (lnMes + 1)) - 1) & "1").Merge
            xlHoja1.Range(ExcelColumnaString(i) & "1:" & ExcelColumnaString(i + (lnCamposNum * (lnMes + 1)) - 1) & "1").HorizontalAlignment = xlCenter
        Next i
        
        xlHoja1.Range(ExcelColumnaString(i) & "1:" & ExcelColumnaString(i + (lnCamposNum) - 1) & "1").Merge
        xlHoja1.Range(ExcelColumnaString(i) & "1:" & ExcelColumnaString(i + (lnCamposNum) - 1) & "1").HorizontalAlignment = xlCenter
        
        
        For i = 3 To lnPosJ - lnCamposNum Step lnCamposNum
            xlHoja1.Range(ExcelColumnaString(i) & "2:" & ExcelColumnaString(i + lnCamposNum - 1) & "2").Merge
            xlHoja1.Range(ExcelColumnaString(i) & "2:" & ExcelColumnaString(i + lnCamposNum - 1) & "2").HorizontalAlignment = xlCenter
        Next i
        
        
        xlHoja1.Range("A1:" & ExcelColumnaString(lnPosJ) & Trim(Str(1))).BorderAround 1, xlMedium
        xlHoja1.Range("A2:" & ExcelColumnaString(lnPosJ) & Trim(Str(2))).BorderAround 1, xlMedium
        xlHoja1.Range("A3:" & ExcelColumnaString(lnPosJ) & Trim(Str(3))).BorderAround 1, xlMedium
        
        For i = 1 To lnPosJ - 1
            xlHoja1.Range(ExcelColumnaString(i) & "1:" & ExcelColumnaString(i) & Trim(Str(lnPosI))).BorderAround 1, xlMedium
        Next i
        
        xlHoja1.Range(ExcelColumnaString(1) & "1:" & ExcelColumnaString(i) & Trim(Str(lnPosI - 1))).BorderAround 1, xlMedium
        xlHoja1.Range(ExcelColumnaString(1) & "1:" & ExcelColumnaString(i) & Trim(Str(lnPosI))).BorderAround 1, xlMedium
        
        xlHoja1.Range("A1:" & ExcelColumnaString(lnPosJ) & Trim(Str(lnPosI))).NumberFormat = "#,##0.00"
        
        xlHoja1.Cells.Select
        xlHoja1.Cells.EntireColumn.AutoFit
    
        xlHoja1.Range("A4:" & ExcelColumnaString(lnPosJ) & Trim(Str(lnPosI - 1))).Select
        'xlHoja1.Range("A4:" & ExcelColumnaString(lnPosJ) & Trim(Str(lnPosI - 1))).Sort xlHoja1.Range(ExcelColumnaString(lnPosJ) & "3"), xlDescending
        
    End If
      
End Sub

'***********************************************************
' Inicia Trabajo con EXCEL, crea variable Aplicacion y Libro
'***********************************************************
Private Function ExcelBegin(psArchivo As String, _
        xlAplicacion As Excel.Application, _
        xlLibro As Excel.Workbook, Optional pbBorraExiste As Boolean = True) As Boolean
        
Dim fs As New Scripting.FileSystemObject
On Error GoTo ErrBegin
Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application

If fs.FileExists(psArchivo) Then
   If pbBorraExiste Then
      fs.DeleteFile psArchivo, True
      Set xlLibro = xlAplicacion.Workbooks.Add
   Else
      Set xlLibro = xlAplicacion.Workbooks.Open(psArchivo)
   End If
Else
   Set xlLibro = xlAplicacion.Workbooks.Add
End If
ExcelBegin = True
Exit Function
ErrBegin:
  MsgBox Err.Description, vbInformation, "Aviso"
  ExcelBegin = False
End Function

'***********************************************************
' Final de Trabajo con EXCEL, graba Libro
'***********************************************************
Private Sub ExcelEnd(psArchivo As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, Optional plSave As Boolean = True)
On Error GoTo ErrEnd
   If plSave Then
        xlHoja1.SaveAs psArchivo
   End If
   xlLibro.Close
   xlAplicacion.Quit
   Set xlAplicacion = Nothing
   Set xlLibro = Nothing
   Set xlHoja1 = Nothing
Exit Sub
ErrEnd:
   MsgBox Err.Description, vbInformation, "Aviso"
End Sub

