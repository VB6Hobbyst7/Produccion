VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPlaPresu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de Presupuesto"
   ClientHeight    =   6780
   ClientLeft      =   120
   ClientTop       =   1275
   ClientWidth     =   11595
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtObj 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   3225
      TabIndex        =   18
      Top             =   660
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   390
      Left            =   10260
      TabIndex        =   17
      Top             =   5790
      Width           =   1170
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   10260
      TabIndex        =   16
      Top             =   6285
      Width           =   1170
   End
   Begin VB.ComboBox cboPresu 
      Height          =   315
      Left            =   3780
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   30
      Width           =   3600
   End
   Begin VB.ComboBox cboFecha 
      Height          =   315
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   30
      Width           =   1000
   End
   Begin VB.TextBox txtObj2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   10170
      TabIndex        =   13
      Top             =   615
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.ComboBox cboMoneda 
      Enabled         =   0   'False
      Height          =   315
      Left            =   8805
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.ComboBox cboTpo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   8700
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   30
      Width           =   1635
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Enabled         =   0   'False
      Height          =   390
      Left            =   10260
      TabIndex        =   10
      Top             =   4830
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   390
      Left            =   10260
      TabIndex        =   9
      Top             =   5295
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Caption         =   "Impresión"
      Height          =   1410
      Left            =   8790
      TabIndex        =   4
      Top             =   5280
      Width           =   1395
      Begin VB.CommandButton cmdImprimir 
         Height          =   390
         Left            =   105
         Picture         =   "frmPlaPresu.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   960
         Width           =   1170
      End
      Begin VB.OptionButton optImp 
         Caption         =   "Mensual"
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   7
         Top             =   210
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.OptionButton optImp 
         Caption         =   "Trimestral"
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   6
         Top             =   450
         Width           =   1020
      End
      Begin VB.OptionButton optImp 
         Caption         =   "Semestral"
         Height          =   285
         Index           =   2
         Left            =   150
         TabIndex        =   5
         Top             =   675
         Width           =   1020
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Distribución"
      Height          =   615
      Left            =   8790
      TabIndex        =   0
      Top             =   4680
      Width           =   1380
      Begin VB.CommandButton cmdMov 
         Enabled         =   0   'False
         Height          =   300
         Left            =   855
         Picture         =   "frmPlaPresu.frx":0D06
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   420
      End
      Begin VB.TextBox txtMes 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   450
         MaxLength       =   2
         TabIndex        =   1
         Text            =   "1"
         Top             =   240
         Width           =   345
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Mes"
         Height          =   210
         Index           =   5
         Left            =   90
         TabIndex        =   3
         Top             =   285
         Width           =   345
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgPresu 
      Height          =   6240
      Left            =   75
      TabIndex        =   19
      Top             =   435
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   11007
      _Version        =   393216
      Cols            =   4
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483638
      TextStyleFixed  =   3
      FocusRect       =   2
      HighLight       =   2
      FillStyle       =   1
      AllowUserResizing=   1
      Appearance      =   0
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgPeriodo 
      Height          =   3900
      Left            =   8760
      TabIndex        =   20
      Top             =   435
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   6879
      _Version        =   393216
      Cols            =   3
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483638
      TextStyleFixed  =   3
      FocusRect       =   2
      HighLight       =   2
      FillStyle       =   1
      AllowUserResizing=   1
      Appearance      =   0
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin RichTextLib.RichTextBox rtfImp 
      Height          =   315
      Left            =   11355
      TabIndex        =   21
      Top             =   4350
      Visible         =   0   'False
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   556
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmPlaPresu.frx":1010
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Nombre :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   2985
      TabIndex        =   27
      Top             =   90
      Width           =   810
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Año :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   825
      TabIndex        =   26
      Top             =   90
      Width           =   495
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10005
      TabIndex        =   25
      Top             =   4305
      Width           =   1305
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   9345
      TabIndex        =   24
      Top             =   4365
      Width           =   600
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Moneda :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   8820
      TabIndex        =   23
      Top             =   465
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Tipo :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   4
      Left            =   8190
      TabIndex        =   22
      Top             =   90
      Width           =   510
   End
End
Attribute VB_Name = "frmPlaPresu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vRTFImp As String
Dim pbIni As Boolean
Dim pRow As Integer

Private Sub cboFecha_Click()
If pbIni Then
    Limpiar
    Call CargaPresu
    Call CargaPeriodo
    Call BloqueoCmd
End If
End Sub
Private Sub cboMoneda_Click()
If pbIni Then
    Limpiar
    Call CargaPresu
    Call CargaPeriodo
End If
End Sub

Private Sub cboPresu_Click()
    Dim tmpSql As String
    Dim sPresu As String
    Dim oPP As DPresupuesto
    Set oPP = New DPresupuesto
    sPresu = Trim(Right(Trim(cboPresu.Text), 4))
    If Len(sPresu) > 0 Then
        tmpSql = oPP.GetPresupuestoTpo(sPresu)
        UbicaCombo cboTpo, tmpSql
    End If
    Call Limpiar
    Call CargaPresu
    Call CargaPeriodo
    Call BloqueoCmd
End Sub

Private Sub CmdCancelar_Click()
'CANCELAR
Call cboPresu_Click
End Sub

Private Sub cmdGrabar_Click()
    'GRABAR
    Dim oPP As DPresupuesto
    Set oPP = New DPresupuesto
    
    Dim nTotalPre As Currency, nTotalMeses As Currency
    Dim nMonto As Currency, nMonCre1 As Currency, nMonCre2 As Currency, nMonIni As Currency
    Dim tmpSql As String
    Dim N As Integer
    txtObj.Visible = False
    txtObj2.Visible = False
    For N = 1 To fgPeriodo.Rows - 1
        nTotalMeses = nTotalMeses + CCur(IIf(fgPeriodo.TextMatrix(N, 3) = "", 0, fgPeriodo.TextMatrix(N, 3)))
    Next
    If Trim(Right(cboTpo.Text, 1)) = "2" Then
        If Not (nTotalMeses = CCur(IIf(Trim(fgPresu.TextMatrix(fgPresu.Row, 9)) = "", 0, Trim(fgPresu.TextMatrix(fgPresu.Row, 9))))) Then
            MsgBox "Falta conciliar cantidades ", vbInformation, " Aviso "
            Exit Sub
        End If
    Else
        If Not (nTotalMeses = CCur(IIf(Trim(fgPresu.TextMatrix(fgPresu.Row, 8)) = "", 0, Trim(fgPresu.TextMatrix(fgPresu.Row, 8))))) Then
            MsgBox "Falta conciliar cantidades ", vbInformation, " Aviso "
            Exit Sub
        End If
    End If
'    If CCur(IIf(Trim(fgPresu.TextMatrix(fgPresu.Row, 8)) = "", 0, Trim(fgPresu.TextMatrix(fgPresu.Row, 8)))) = 0 And _
'    CCur(IIf(Trim(fgPresu.TextMatrix(fgPresu.Row, 4)) = "", 0, Trim(fgPresu.TextMatrix(fgPresu.Row, 4)))) = 0 Then
'        MsgBox "Determine los montos", vbInformation, " Aviso"
'        Exit Sub
'    End If
    nMonIni = CCur(IIf(Trim(fgPresu.TextMatrix(fgPresu.Row, 4)) = "", 0, Trim(fgPresu.TextMatrix(fgPresu.Row, 4))))
    nMonto = CCur(IIf(Trim(fgPresu.TextMatrix(fgPresu.Row, 5)) = "", 0, Trim(fgPresu.TextMatrix(fgPresu.Row, 5))))
    nMonCre1 = CCur(IIf(Trim(fgPresu.TextMatrix(fgPresu.Row, 6)) = "", 0, Trim(fgPresu.TextMatrix(fgPresu.Row, 6))))
    nMonCre2 = CCur(IIf(Trim(fgPresu.TextMatrix(fgPresu.Row, 7)) = "", 0, Trim(fgPresu.TextMatrix(fgPresu.Row, 7))))
    'BALANCE
    If Trim(Right(cboTpo.Text, 1)) = "2" Then
        nTotalPre = CCur(IIf(Trim(fgPresu.TextMatrix(fgPresu.Row, 9)) = "", 0, Trim(fgPresu.TextMatrix(fgPresu.Row, 9))))
    Else
        nTotalPre = nMonto + nMonCre1 + nMonCre2
    End If
    If nTotalMeses <> nTotalPre Then
        MsgBox "Cantidades de periodos no coinciden con monto total", vbInformation, " Aviso "
        Exit Sub
    End If
    If fgPresu.TextMatrix(fgPresu.Row, 3) = "G" Then
        MsgBox "Ya se grabó la información", vbInformation, " Aviso "
        Exit Sub
    End If
    If MsgBox("Esta seguro de Grabar estas cantidades", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
        oPP.ModificaMontoRubro Right(Me.cboPresu, 4), Me.cboFecha.Text, fgPresu.TextMatrix(fgPresu.Row, 1), nMonIni, nMonto, nMonCre1, nMonCre2
        For N = 1 To fgPeriodo.Rows - 1
            If fgPresu.TextMatrix(fgPresu.Row, 3) = "M" Then
                oPP.ModificaMontoRubroMes Right(Me.cboPresu, 4), Me.cboFecha.Text, fgPresu.TextMatrix(fgPresu.Row, 1), fgPeriodo.TextMatrix(N, 1), CCur(IIf(fgPeriodo.TextMatrix(N, 3) = "", 0, fgPeriodo.TextMatrix(N, 3)))
            Else
                oPP.AgregaMontoRubroMes Right(Me.cboPresu, 4), Me.cboFecha.Text, fgPresu.TextMatrix(fgPresu.Row, 1), fgPeriodo.TextMatrix(N, 1), CCur(IIf(fgPeriodo.TextMatrix(N, 3) = "", 0, fgPeriodo.TextMatrix(N, 3)))
            End If
        Next
        fgPresu.TextMatrix(fgPresu.Row, 3) = "G"
        Call fgPresu_RowColChange
    Else
        MsgBox "Grabación Cancelada", vbInformation, " Aviso "
    End If
End Sub

Private Sub cmdImprimir_Click()
Dim sArchivo As String
MousePointer = 11

sArchivo = "PresupuestoRub_" & Me.cboFecha & ".XLS"

Call GeneraReporte(sArchivo, _
    Switch(optImp(0).value = True, 1, optImp(1).value = True, 2, optImp(2).value = True, 3))
If sArchivo <> "" Then
    '*******Carga el Archivo Excel a Objeto Ole ******
    CargaArchivo sArchivo, App.path & "\SPOOLER\"
End If
MousePointer = 0
End Sub

Private Sub GeneraReporte(ByVal psArchivo As String, ByVal pnTipo As Integer)
On Error GoTo ErrorGeneraReporte
    Dim fs As New Scripting.FileSystemObject
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExcel As Boolean
    Dim nTotal(25)
    Dim nMonBal As Currency
    Dim psNomHoja As String
    
    Dim tmpReg As ADODB.Recordset
    Set tmpReg = New ADODB.Recordset
    Dim oPP As DPresupuesto
    Set oPP = New DPresupuesto
    Dim tmpSql As String
    Dim nFil As Integer, nCol As Integer, nItem As Integer
    
    Set fs = New Scripting.FileSystemObject
    Set xlAplicacion = New Excel.Application
    
    If fs.FileExists(App.path & "\SPOOLER\" & psArchivo) Then
        Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & psArchivo)
    Else
        Set xlLibro = xlAplicacion.Workbooks.Add
    End If
    lbExcel = True
    Select Case pnTipo
      Case 1:  psNomHoja = "MENSUAL"
      Case 2:  psNomHoja = "TRIMESTRAL"
      Case 3:  psNomHoja = "SEMESTRAL"
   End Select
    For Each xlHoja1 In xlLibro.Worksheets
        If xlHoja1.Name = psNomHoja Then
            xlHoja1.Delete
            Exit For
        End If
    Next
    
    Set xlHoja1 = xlLibro.Worksheets.Add
    xlHoja1.Name = psNomHoja
    
    xlHoja1.PageSetup.Zoom = 70
    xlAplicacion.Range("A1:W100").Font.Size = 7
    
    xlHoja1.Range("A1").ColumnWidth = 5
    xlHoja1.Range("B1").ColumnWidth = 10
    xlHoja1.Range("C1").ColumnWidth = 20
    xlHoja1.Range("D1..K1").ColumnWidth = 11
    xlHoja1.Range("L1..W1").ColumnWidth = 10
    
    xlHoja1.Cells(1, 1) = gsNomCmac
    xlAplicacion.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 1)).Font.Bold = True
    If pnTipo = 1 Then
        xlHoja1.Cells(1, 21) = "Fecha :" & Format(gdFecSis, "dd mmmm yyyy")
        xlAplicacion.Range(xlHoja1.Cells(1, 21), xlHoja1.Cells(1, 21)).Font.Bold = True
    ElseIf pnTipo = 2 Then
        xlHoja1.Cells(1, 13) = "Fecha :" & Format(gdFecSis, "dd mmmm yyyy")
        xlAplicacion.Range(xlHoja1.Cells(1, 13), xlHoja1.Cells(1, 13)).Font.Bold = True
    Else
        xlHoja1.Cells(1, 10) = "Fecha :" & Format(gdFecSis, "dd mmmm yyyy")
        xlAplicacion.Range(xlHoja1.Cells(1, 10), xlHoja1.Cells(1, 10)).Font.Bold = True
    End If
    
    xlHoja1.Cells(2, 1) = "Area de Planeamiento"
    xlAplicacion.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 1)).Font.Bold = True
    
    If pnTipo = 1 Then
        xlHoja1.Cells(3, 12) = Trim(Mid(cboPresu.Text, 1, Len(cboPresu.Text) - 4))
        xlAplicacion.Range(xlHoja1.Cells(3, 12), xlHoja1.Cells(3, 12)).Font.Bold = True
        
        xlHoja1.Cells(4, 13) = " AÑO : " & Trim(cboFecha.Text)
        xlAplicacion.Range(xlHoja1.Cells(4, 13), xlHoja1.Cells(4, 13)).Font.Bold = True
    ElseIf pnTipo = 2 Then
        xlHoja1.Cells(3, 7) = Trim(Mid(cboPresu.Text, 1, Len(cboPresu.Text) - 4))
        xlAplicacion.Range(xlHoja1.Cells(3, 7), xlHoja1.Cells(3, 7)).Font.Bold = True
        
        xlHoja1.Cells(4, 8) = " AÑO : " & Trim(cboFecha.Text)
        xlAplicacion.Range(xlHoja1.Cells(4, 8), xlHoja1.Cells(4, 8)).Font.Bold = True
    Else
        xlHoja1.Cells(3, 4) = Trim(Mid(cboPresu.Text, 1, Len(cboPresu.Text) - 4))
        xlAplicacion.Range(xlHoja1.Cells(3, 4), xlHoja1.Cells(3, 4)).Font.Bold = True
        
        xlHoja1.Cells(4, 5) = " AÑO : " & Trim(cboFecha.Text)
        xlAplicacion.Range(xlHoja1.Cells(4, 5), xlHoja1.Cells(4, 5)).Font.Bold = True
    End If
    
    xlAplicacion.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(7, 22)).HorizontalAlignment = xlHAlignCenter
    xlAplicacion.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(7, 22)).Font.Bold = True
    nCol = 1
    xlHoja1.Cells(7, 1) = "ITEM": xlHoja1.Cells(7, 2) = "CODIGO": xlHoja1.Cells(7, 3) = "DESCRIPCION": xlHoja1.Cells(7, 4) = "AÑO " & Val(Trim(cboFecha.Text)) - 1: xlHoja1.Cells(7, 5) = "PRESUPUESTO": xlHoja1.Cells(7, 6) = "CRED.1": xlHoja1.Cells(7, 7) = "CRED.2": xlHoja1.Cells(7, 8) = "TOTAL": xlHoja1.Cells(7, 9) = "Var.Monto": xlHoja1.Cells(7, 10) = "Var. %"
    
    CuadroExcel xlHoja1, 1, 6, 10, 7
    If pnTipo = 1 Then
        CuadroExcel xlHoja1, 11, 6, 22, 6, True
        CuadroExcel xlHoja1, 11, 7, 22, 7
        xlHoja1.Cells(6, 15) = "M   E   S   E   S "
        xlHoja1.Cells(7, 11) = "ENERO": xlHoja1.Cells(7, 12) = "FEBRERO": xlHoja1.Cells(7, 13) = "MARZO": xlHoja1.Cells(7, 14) = "ABRIL": xlHoja1.Cells(7, 15) = "MAYO": xlHoja1.Cells(7, 16) = "JUNIO": xlHoja1.Cells(7, 17) = "JULIO": xlHoja1.Cells(7, 18) = "AGOSTO": xlHoja1.Cells(7, 19) = "SETIEMBRE": xlHoja1.Cells(7, 20) = "OCTUBRE": xlHoja1.Cells(7, 21) = "NOVIEMBRE": xlHoja1.Cells(7, 22) = "DICIEMBRE"
    ElseIf pnTipo = 2 Then
        CuadroExcel xlHoja1, 11, 6, 14, 6, True
        CuadroExcel xlHoja1, 11, 7, 14, 7
        xlHoja1.Cells(6, 12) = "T R I M E S T R E S"
        xlHoja1.Cells(7, 11) = "PRIMERO": xlHoja1.Cells(7, 12) = "SEGUNDO": xlHoja1.Cells(7, 13) = "TERCERO": xlHoja1.Cells(7, 14) = "CUARTO"
    Else
        CuadroExcel xlHoja1, 11, 6, 12, 6, True
        CuadroExcel xlHoja1, 11, 7, 12, 7
        xlHoja1.Cells(6, 11) = "S E M E S T R E S"
        xlHoja1.Cells(7, 11) = "PRIMERO": xlHoja1.Cells(7, 12) = "SEGUNDO"
    End If
    nFil = 7
    nItem = 0

    Set tmpReg = oPP.GetPresupTipRep(pnTipo, Me.cboFecha.Text, Right(cboPresu.Text, 4))
    If Not (tmpReg.BOF Or tmpReg.EOF) Then
        With tmpReg
            nTotal(2) = !cCodRub
            Do While Not .EOF
                'SOLO BALANCES
                If Right(Trim(cboTpo.Text), 1) = "2" Then
                    If Len(Trim(!cCodRub)) = 4 Then
                        If nTotal(2) <> !cCodRub Then
                            nFil = nFil + 1
                            xlHoja1.Cells(nFil, 2) = ""
                            xlHoja1.Cells(nFil, 3) = "TOTAL  " & nTotal(3)
                            xlHoja1.Cells(nFil, 4) = nTotal(4)
                            xlHoja1.Cells(nFil, 5) = nTotal(5)
                            xlHoja1.Cells(nFil, 4) = nTotal(4)
                            xlHoja1.Cells(nFil, 5) = nTotal(5)
                            xlHoja1.Cells(nFil, 6) = nTotal(6)
                            xlHoja1.Cells(nFil, 7) = nTotal(7)
                            xlHoja1.Cells(nFil, 8) = nTotal(8)
                            xlHoja1.Cells(nFil, 9) = nTotal(9)
                            xlHoja1.Cells(nFil, 10) = nTotal(10)
                            If pnTipo = 1 Then
                                xlHoja1.Cells(nFil, 11) = nTotal(11)
                                xlHoja1.Cells(nFil, 12) = nTotal(12)
                                xlHoja1.Cells(nFil, 13) = nTotal(13)
                                xlHoja1.Cells(nFil, 14) = nTotal(14)
                                xlHoja1.Cells(nFil, 15) = nTotal(15)
                                xlHoja1.Cells(nFil, 16) = nTotal(16)
                                xlHoja1.Cells(nFil, 17) = nTotal(17)
                                xlHoja1.Cells(nFil, 18) = nTotal(18)
                                xlHoja1.Cells(nFil, 19) = nTotal(19)
                                xlHoja1.Cells(nFil, 20) = nTotal(20)
                                xlHoja1.Cells(nFil, 21) = nTotal(21)
                                xlHoja1.Cells(nFil, 22) = nTotal(22)
                            ElseIf pnTipo = 2 Then
                                xlHoja1.Cells(nFil, 11) = nTotal(11)
                                xlHoja1.Cells(nFil, 12) = nTotal(12)
                                xlHoja1.Cells(nFil, 13) = nTotal(13)
                                xlHoja1.Cells(nFil, 14) = nTotal(14)
                            Else
                                xlHoja1.Cells(nFil, 11) = nTotal(11)
                                xlHoja1.Cells(nFil, 12) = nTotal(12)
                            End If
                            nFil = nFil + 1
                        End If
                        
                        nTotal(2) = !cCodRub
                        nTotal(3) = !cDesRub
                        If !nMonIni <> 0 Or !nMonto <> 0 Then
                            nTotal(4) = Format(!nMonIni, "#,##0.00")
                            nTotal(5) = Format(!nMonto, "#,##0.00")
                        Else
                            nTotal(4) = IIf(!nMonIni = 0, "", Format(!nMonIni, "#,##0.00"))
                            nTotal(5) = IIf(!nMonto = 0, "", Format(!nMonto, "#,##0.00"))
                        End If
                        nTotal(6) = IIf(!nMonCre1 = 0, "", Format(!nMonCre1, "#,##0.00"))
                        nTotal(7) = IIf(!nMonCre2 = 0, "", Format(!nMonCre2, "#,##0.00"))
                        If !nMonIni <> 0 Or !nMonto <> 0 Then
                            nTotal(8) = Format(!Total, "#,##0.00")
                        Else
                            nTotal(8) = IIf(!Total = 0, "", Format(!Total, "#,##0.00"))
                        End If
                        If !Total <> 0 And !nMonIni <> 0 Then
                            nTotal(9) = Format(!Total - !nMonIni, "#,##0.00")
                            nTotal(10) = Format((((!Total / !nMonIni) - 1) * 100), "#,##0.00")
                        ElseIf !Total <> 0 Or !nMonIni <> 0 Then
                            nTotal(9) = Format(!Total - !nMonIni, "#,##0.00")
                            If !Total = 0 Then
                                nTotal(10) = Format(100 * -1, "#,##0.00")
                            Else
                                nTotal(10) = Format(100, "#,##0.00")
                            End If
                        Else
                            nTotal(9) = ""
                            nTotal(10) = ""
                        End If
                        If pnTipo = 1 Then
                            If !nMonIni <> 0 Or !nMonto <> 0 Then
                                nMonBal = !nMonIni
                                nMonBal = nMonBal + !Enero
                                nTotal(11) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Febrero
                                nTotal(12) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Marzo
                                nTotal(13) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Abril
                                nTotal(14) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Mayo
                                nTotal(15) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Junio
                                nTotal(16) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Julio
                                nTotal(17) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Agosto
                                nTotal(18) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Setiembre
                                nTotal(19) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Octubre
                                nTotal(20) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Noviembre
                                nTotal(21) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Diciembre
                                nTotal(22) = Format(nMonBal, "#,##0.00")
                            Else
                                nMonBal = IIf(IsNull(!nMonIni), 0, !nMonIni)
                                nMonBal = nMonBal + IIf(IsNull(!Enero), 0, !Enero)
                                nTotal(11) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + IIf(IsNull(!Febrero), 0, !Febrero)
                                nTotal(12) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + IIf(IsNull(!Marzo), 0, !Marzo)
                                nTotal(13) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + IIf(IsNull(!Abril), 0, !Abril)
                                nTotal(14) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + IIf(IsNull(!Mayo), 0, !Mayo)
                                nTotal(15) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + IIf(IsNull(!Junio), 0, !Junio)
                                nTotal(16) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + IIf(IsNull(!Julio), 0, !Julio)
                                nTotal(17) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + IIf(IsNull(!Agosto), 0, !Agosto)
                                nTotal(18) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + IIf(IsNull(!Setiembre), 0, !Enero)
                                nTotal(19) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + IIf(IsNull(!Octubre), 0, !Octubre)
                                nTotal(20) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + IIf(IsNull(!Noviembre), 0, !Noviembre)
                                nTotal(21) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + IIf(IsNull(!Diciembre), 0, !Diciembre)
                                nTotal(22) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                            End If
                        ElseIf pnTipo = 2 Then
                            If !nMonIni <> 0 Or !nMonto <> 0 Then
                                nMonBal = !nMonIni
                                nMonBal = nMonBal + !Primero
                                nTotal(11) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Segundo
                                nTotal(12) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Tercero
                                nTotal(13) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Cuarto
                                nTotal(14) = Format(nMonBal, "#,##0.00")
                            Else
                                nMonBal = !nMonIni
                                nMonBal = nMonBal + !Primero
                                nTotal(11) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + !Segundo
                                nTotal(12) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + !Tercero
                                nTotal(13) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + !Cuarto
                                nTotal(14) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                            End If
                        Else
                            If !nMonIni <> 0 Or !nMonto <> 0 Then
                                nMonBal = !nMonIni
                                nMonBal = nMonBal + !Primero
                                nTotal(11) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Segundo
                                nTotal(12) = Format(nMonBal, "#,##0.00")
                            Else
                                nMonBal = !nMonIni
                                nMonBal = nMonBal + !Primero
                                nTotal(11) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + !Segundo
                                nTotal(12) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                            End If
                        End If
                    End If
                    
                    nFil = nFil + 1
                    nItem = nItem + 1
                    xlHoja1.Cells(nFil, 1) = nItem
                    xlHoja1.Cells(nFil, 2) = !cCodRub
                    xlHoja1.Cells(nFil, 3) = !cDesRub
                    If Len(!cCodRub) <> 4 Then
                        If !nMonIni <> 0 Or !nMonto <> 0 Then
                            xlHoja1.Cells(nFil, 4) = Format(!nMonIni, "#,##0.00")
                            xlHoja1.Cells(nFil, 5) = Format(!nMonto, "#,##0.00")
                        Else
                            xlHoja1.Cells(nFil, 4) = IIf(!nMonIni = 0, "", Format(!nMonIni, "#,##0.00"))
                            xlHoja1.Cells(nFil, 5) = IIf(!nMonto = 0, "", Format(!nMonto, "#,##0.00"))
                        End If
                        xlHoja1.Cells(nFil, 6) = IIf(!nMonCre1 = 0, "", Format(!nMonCre1, "#,##0.00"))
                        xlHoja1.Cells(nFil, 7) = IIf(!nMonCre2 = 0, "", Format(!nMonCre2, "#,##0.00"))
                        If !nMonIni <> 0 Or !nMonto <> 0 Then
                            xlHoja1.Cells(nFil, 8) = Format(!Total, "#,##0.00")
                        Else
                            xlHoja1.Cells(nFil, 8) = IIf(!Total = 0, "", Format(!Total, "#,##0.00"))
                        End If
                        If !Total <> 0 And !nMonIni <> 0 Then
                            xlHoja1.Cells(nFil, 9) = Format(!Total - !nMonIni, "#,##0.00")
                            xlHoja1.Cells(nFil, 10) = Format((((!Total / !nMonIni) - 1) * 100), "#,##0.00")
                        ElseIf !Total <> 0 Or !nMonIni <> 0 Then
                            xlHoja1.Cells(nFil, 9) = Format(!Total - !nMonIni, "#,##0.00")
                            If !Total = 0 Then
                                xlHoja1.Cells(nFil, 10) = Format(100 * -1, "#,##0.00")
                            Else
                                xlHoja1.Cells(nFil, 10) = Format(100, "#,##0.00")
                            End If
                        Else
                            xlHoja1.Cells(nFil, 9) = ""
                            xlHoja1.Cells(nFil, 10) = ""
                        End If
                        If pnTipo = 1 Then
                            If !nMonIni <> 0 Or !nMonto <> 0 Then
                                nMonBal = !nMonIni
                                nMonBal = nMonBal + !Enero
                                xlHoja1.Cells(nFil, 11) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Febrero
                                xlHoja1.Cells(nFil, 12) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Marzo
                                xlHoja1.Cells(nFil, 13) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Abril
                                xlHoja1.Cells(nFil, 14) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Mayo
                                xlHoja1.Cells(nFil, 15) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Junio
                                xlHoja1.Cells(nFil, 16) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Julio
                                xlHoja1.Cells(nFil, 17) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Agosto
                                xlHoja1.Cells(nFil, 18) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Setiembre
                                xlHoja1.Cells(nFil, 19) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Octubre
                                xlHoja1.Cells(nFil, 20) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Noviembre
                                xlHoja1.Cells(nFil, 21) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Diciembre
                                xlHoja1.Cells(nFil, 22) = Format(nMonBal, "#,##0.00")
                            Else
                                nMonBal = IIf(IsNull(!nMonIni), 0, !nMonIni)
                                nMonBal = nMonBal + IIf(IsNull(!Enero), 0, !nMonIni)
                                xlHoja1.Cells(nFil, 11) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + IIf(IsNull(!Febrero), 0, !Febrero)
                                xlHoja1.Cells(nFil, 12) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + IIf(IsNull(!Marzo), 0, !Marzo)
                                xlHoja1.Cells(nFil, 13) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + IIf(IsNull(!Abril), 0, !Abril)
                                xlHoja1.Cells(nFil, 14) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + IIf(IsNull(!Mayo), 0, !Mayo)
                                xlHoja1.Cells(nFil, 15) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + IIf(IsNull(!Junio), 0, !Junio)
                                xlHoja1.Cells(nFil, 16) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + IIf(IsNull(!Julio), 0, !Julio)
                                xlHoja1.Cells(nFil, 17) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + IIf(IsNull(!Agosto), 0, !Agosto)
                                xlHoja1.Cells(nFil, 18) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + IIf(IsNull(!Setiembre), 0, !Setiembre)
                                xlHoja1.Cells(nFil, 19) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + IIf(IsNull(!Octubre), 0, !Octubre)
                                xlHoja1.Cells(nFil, 20) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + IIf(IsNull(!Noviembre), 0, !Noviembre)
                                xlHoja1.Cells(nFil, 21) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + IIf(IsNull(!Diciembre), 0, !Diciembre)
                                xlHoja1.Cells(nFil, 22) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                            End If
                        ElseIf pnTipo = 2 Then
                            If !nMonIni <> 0 Or !nMonto <> 0 Then
                                nMonBal = !nMonIni
                                nMonBal = nMonBal + !Primero
                                xlHoja1.Cells(nFil, 11) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Segundo
                                xlHoja1.Cells(nFil, 12) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Tercero
                                xlHoja1.Cells(nFil, 13) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Cuarto
                                xlHoja1.Cells(nFil, 14) = Format(nMonBal, "#,##0.00")
                            Else
                                nMonBal = !nMonIni
                                nMonBal = nMonBal + !Primero
                                xlHoja1.Cells(nFil, 11) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + !Segundo
                                xlHoja1.Cells(nFil, 12) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + !Tercero
                                xlHoja1.Cells(nFil, 13) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + !Cuarto
                                xlHoja1.Cells(nFil, 14) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                            End If
                        Else
                            If !nMonIni <> 0 Or !nMonto <> 0 Then
                                nMonBal = !nMonIni
                                nMonBal = nMonBal + !Primero
                                xlHoja1.Cells(nFil, 11) = Format(nMonBal, "#,##0.00")
                                nMonBal = nMonBal + !Segundo
                                xlHoja1.Cells(nFil, 12) = Format(nMonBal, "#,##0.00")
                            Else
                                nMonBal = !nMonIni
                                nMonBal = nMonBal + !Primero
                                xlHoja1.Cells(nFil, 11) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                                nMonBal = nMonBal + !Segundo
                                xlHoja1.Cells(nFil, 12) = IIf(nMonBal = 0, "", Format(nMonBal, "#,##0.00"))
                            End If
                        End If
                    End If
                Else
                    'OTROS (NO BALANCES)
                    nFil = nFil + 1
                    nItem = nItem + 1
                    xlHoja1.Cells(nFil, 1) = nItem
                    xlHoja1.Cells(nFil, 2) = !cCodRub
                    xlHoja1.Cells(nFil, 3) = !cDesRub
                    If !nMonIni <> 0 Or !nMonto <> 0 Then
                        xlHoja1.Cells(nFil, 4) = Format(!nMonIni, "#,##0.00")
                        xlHoja1.Cells(nFil, 5) = Format(!nMonto, "#,##0.00")
                    Else
                        xlHoja1.Cells(nFil, 4) = IIf(!nMonIni = 0, "", Format(!nMonIni, "#,##0.00"))
                        xlHoja1.Cells(nFil, 5) = IIf(!nMonto = 0, "", Format(!nMonto, "#,##0.00"))
                    End If
                    xlHoja1.Cells(nFil, 6) = IIf(!nMonCre1 = 0, "", Format(!nMonCre1, "#,##0.00"))
                    xlHoja1.Cells(nFil, 7) = IIf(!nMonCre2 = 0, "", Format(!nMonCre2, "#,##0.00"))
                    If !nMonIni <> 0 Or !nMonto <> 0 Then
                        xlHoja1.Cells(nFil, 8) = Format(!Total, "#,##0.00")
                    Else
                        xlHoja1.Cells(nFil, 8) = IIf(!Total = 0, "", Format(!Total, "#,##0.00"))
                    End If
                    If !Total <> 0 And !nMonIni <> 0 Then
                        xlHoja1.Cells(nFil, 9) = Format(!Total - !nMonIni, "#,##0.00")
                        xlHoja1.Cells(nFil, 10) = Format((((!Total / !nMonIni) - 1) * 100), "#,##0.00")
                    ElseIf !Total <> 0 Or !nMonIni <> 0 Then
                        xlHoja1.Cells(nFil, 9) = Format(!Total - !nMonIni, "#,##0.00")
                        If !Total = 0 Then
                            xlHoja1.Cells(nFil, 10) = Format(100 * -1, "#,##0.00")
                        Else
                            xlHoja1.Cells(nFil, 10) = Format(100, "#,##0.00")
                        End If
                    Else
                        xlHoja1.Cells(nFil, 9) = ""
                        xlHoja1.Cells(nFil, 10) = ""
                    End If
                    If pnTipo = 1 Then
                        If !nMonIni <> 0 Or !nMonto <> 0 Then
                            xlHoja1.Cells(nFil, 11) = Format(!Enero, "#,##0.00")
                            xlHoja1.Cells(nFil, 12) = Format(!Febrero, "#,##0.00")
                            xlHoja1.Cells(nFil, 13) = Format(!Marzo, "#,##0.00")
                            xlHoja1.Cells(nFil, 14) = Format(!Abril, "#,##0.00")
                            xlHoja1.Cells(nFil, 15) = Format(!Mayo, "#,##0.00")
                            xlHoja1.Cells(nFil, 16) = Format(!Junio, "#,##0.00")
                            xlHoja1.Cells(nFil, 17) = Format(!Julio, "#,##0.00")
                            xlHoja1.Cells(nFil, 18) = Format(!Agosto, "#,##0.00")
                            xlHoja1.Cells(nFil, 19) = Format(!Setiembre, "#,##0.00")
                            xlHoja1.Cells(nFil, 20) = Format(!Octubre, "#,##0.00")
                            xlHoja1.Cells(nFil, 21) = Format(!Noviembre, "#,##0.00")
                            xlHoja1.Cells(nFil, 22) = Format(!Diciembre, "#,##0.00")
                        Else
                            xlHoja1.Cells(nFil, 11) = IIf(!Enero = 0, "", Format(!Enero, "#,##0.00"))
                            xlHoja1.Cells(nFil, 12) = IIf(!Febrero = 0, "", Format(!Febrero, "#,##0.00"))
                            xlHoja1.Cells(nFil, 13) = IIf(!Marzo = 0, "", Format(!Marzo, "#,##0.00"))
                            xlHoja1.Cells(nFil, 14) = IIf(!Abril = 0, "", Format(!Abril, "#,##0.00"))
                            xlHoja1.Cells(nFil, 15) = IIf(!Mayo = 0, "", Format(!Mayo, "#,##0.00"))
                            xlHoja1.Cells(nFil, 16) = IIf(!Junio = 0, "", Format(!Junio, "#,##0.00"))
                            xlHoja1.Cells(nFil, 17) = IIf(!Julio = 0, "", Format(!Julio, "#,##0.00"))
                            xlHoja1.Cells(nFil, 18) = IIf(!Agosto = 0, "", Format(!Agosto, "#,##0.00"))
                            xlHoja1.Cells(nFil, 19) = IIf(!Setiembre = 0, "", Format(!Setiembre, "#,##0.00"))
                            xlHoja1.Cells(nFil, 20) = IIf(!Octubre = 0, "", Format(!Octubre, "#,##0.00"))
                            xlHoja1.Cells(nFil, 21) = IIf(!Noviembre = 0, "", Format(!Noviembre, "#,##0.00"))
                            xlHoja1.Cells(nFil, 22) = IIf(!Diciembre = 0, "", Format(!Diciembre, "#,##0.00"))
                        End If
                    ElseIf pnTipo = 2 Then
                        If !nMonIni <> 0 Or !nMonto <> 0 Then
                            xlHoja1.Cells(nFil, 11) = Format(!Primero, "#,##0.00")
                            xlHoja1.Cells(nFil, 12) = Format(!Segundo, "#,##0.00")
                            xlHoja1.Cells(nFil, 13) = Format(!Tercero, "#,##0.00")
                            xlHoja1.Cells(nFil, 14) = Format(!Cuarto, "#,##0.00")
                        Else
                            xlHoja1.Cells(nFil, 11) = IIf(!Primero = 0, "", Format(!Primero, "#,##0.00"))
                            xlHoja1.Cells(nFil, 12) = IIf(!Segundo = 0, "", Format(!Segundo, "#,##0.00"))
                            xlHoja1.Cells(nFil, 13) = IIf(!Tercero = 0, "", Format(!Tercero, "#,##0.00"))
                            xlHoja1.Cells(nFil, 14) = IIf(!Cuarto = 0, "", Format(!Cuarto, "#,##0.00"))
                        End If
                    Else
                        If !nMonIni <> 0 Or !nMonto <> 0 Then
                            xlHoja1.Cells(nFil, 11) = Format(!Primero, "#,##0.00")
                            xlHoja1.Cells(nFil, 12) = Format(!Segundo, "#,##0.00")
                        Else
                            xlHoja1.Cells(nFil, 11) = IIf(!Primero = 0, "", Format(!Primero, "#,##0.00"))
                            xlHoja1.Cells(nFil, 12) = IIf(!Segundo = 0, "", Format(!Segundo, "#,##0.00"))
                        End If
                    End If
                End If
                .MoveNext
            Loop
        End With
        'SOLO BALANCES
        If Right(Trim(cboTpo.Text), 1) = "2" Then
            'Lista el último TOTAL
            nFil = nFil + 1
            xlHoja1.Cells(nFil, 2) = ""
            xlHoja1.Cells(nFil, 3) = "TOTAL  " & nTotal(3)
            xlHoja1.Cells(nFil, 4) = nTotal(4)
            xlHoja1.Cells(nFil, 5) = nTotal(5)
            xlHoja1.Cells(nFil, 4) = nTotal(4)
            xlHoja1.Cells(nFil, 5) = nTotal(5)
            xlHoja1.Cells(nFil, 6) = nTotal(6)
            xlHoja1.Cells(nFil, 7) = nTotal(7)
            xlHoja1.Cells(nFil, 8) = nTotal(8)
            xlHoja1.Cells(nFil, 9) = nTotal(9)
            xlHoja1.Cells(nFil, 10) = nTotal(10)
            If pnTipo = 1 Then
                xlHoja1.Cells(nFil, 11) = nTotal(11)
                xlHoja1.Cells(nFil, 12) = nTotal(12)
                xlHoja1.Cells(nFil, 13) = nTotal(13)
                xlHoja1.Cells(nFil, 14) = nTotal(14)
                xlHoja1.Cells(nFil, 15) = nTotal(15)
                xlHoja1.Cells(nFil, 16) = nTotal(16)
                xlHoja1.Cells(nFil, 17) = nTotal(17)
                xlHoja1.Cells(nFil, 18) = nTotal(18)
                xlHoja1.Cells(nFil, 19) = nTotal(19)
                xlHoja1.Cells(nFil, 20) = nTotal(20)
                xlHoja1.Cells(nFil, 21) = nTotal(21)
                xlHoja1.Cells(nFil, 22) = nTotal(22)
            ElseIf pnTipo = 2 Then
                xlHoja1.Cells(nFil, 11) = nTotal(11)
                xlHoja1.Cells(nFil, 12) = nTotal(12)
                xlHoja1.Cells(nFil, 13) = nTotal(13)
                xlHoja1.Cells(nFil, 14) = nTotal(14)
            Else
                xlHoja1.Cells(nFil, 11) = nTotal(11)
                xlHoja1.Cells(nFil, 12) = nTotal(12)
            End If
        End If
    End If
    tmpReg.Close
    Set tmpReg = Nothing
    If pnTipo = 1 Then
        CuadroExcel xlHoja1, 1, 8, 22, nFil
    ElseIf pnTipo = 2 Then
        CuadroExcel xlHoja1, 1, 8, 14, nFil
    Else
        CuadroExcel xlHoja1, 1, 8, 12, nFil
    End If
    
    xlHoja1.SaveAs App.path & "\SPOOLER\" & psArchivo
    xlLibro.Close
    xlAplicacion.Quit
    Set xlAplicacion = Nothing
    Set xlLibro = Nothing
    Set xlHoja1 = Nothing
    lbExcel = False
    
Exit Sub
ErrorGeneraReporte:
    MousePointer = 0
    MsgBox "Error Nº [" & Err.Number & " ] " & Err.Description, vbInformation, "Aviso"
    If lbExcel = True Then
        xlLibro.Close
        xlAplicacion.Quit
        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing
    End If
End Sub

Private Sub CuadroExcel(plHoja1 As Excel.Worksheet, X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer, Optional lbLineasVert As Boolean = False)
Dim i, J As Integer

For i = X1 To X2
    plHoja1.Range(plHoja1.Cells(Y1, i), plHoja1.Cells(Y1, i)).Borders(xlEdgeTop).LineStyle = xlContinuous
    plHoja1.Range(plHoja1.Cells(Y2, i), plHoja1.Cells(Y2, i)).Borders(xlEdgeBottom).LineStyle = xlContinuous
Next i
If lbLineasVert = False Then
    For i = X1 To X2
        For J = Y1 To Y2
            plHoja1.Range(plHoja1.Cells(J, i), plHoja1.Cells(J, i)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Next J
    Next i
End If
If lbLineasVert Then
    For J = Y1 To Y2
        plHoja1.Range(plHoja1.Cells(J, X1), plHoja1.Cells(J, X1)).Borders(xlEdgeRight).LineStyle = xlContinuous
    Next J
End If

For J = Y1 To Y2
    plHoja1.Range(plHoja1.Cells(J, X2), plHoja1.Cells(J, X2)).Borders(xlEdgeRight).LineStyle = xlContinuous
Next J
End Sub

Private Sub cmdModificar_Click()
'MODIFICAR
If fgPresu.TextMatrix(fgPresu.Row, 3) = "G" Then
    pRow = fgPresu.Row
    cmdModificar.Enabled = False
    cmdGrabar.Enabled = True
    cmdCancelar.Enabled = True
    cmdMov.Enabled = True
    fgPresu.TextMatrix(fgPresu.Row, 3) = "M"
    
End If
End Sub

Private Sub cmdMov_Click()
'MOV - R = r+r+...+r
Dim N As Integer
Dim nMesIni As Integer
Dim nTotalPre As Currency, nTotalMes As Currency, nTotalMeses As Currency
Dim nTotalMesAnt As Currency
'Valida Meses
nMesIni = Val(txtMes.Text)
If Not (nMesIni > 0 And nMesIni <= 12) Then
    MsgBox "Ingrese el número de mes desde donde se iniciará la distribución", vbInformation, " Aviso "
    txtMes.Text = ""
    Exit Sub
End If
If Trim(Right(cboTpo.Text, 1)) = "2" Then
    nTotalPre = CCur(IIf(Trim(fgPresu.TextMatrix(fgPresu.Row, 9)) = "", 0, Trim(fgPresu.TextMatrix(fgPresu.Row, 9))))
Else
    nTotalPre = CCur(IIf(Trim(fgPresu.TextMatrix(fgPresu.Row, 8)) = "", 0, Trim(fgPresu.TextMatrix(fgPresu.Row, 8))))
End If
'nTotalPre <> 0 And
If (fgPresu.Col >= 4 And fgPresu.Col <= 7) And _
    IsUltNivel(Left(Trim(cboFecha.Text), 4), Right(cboPresu.Text, 4), fgPresu.TextMatrix(fgPresu.Row, 1)) Then
    txtObj2.Visible = False
    'fgPresu.TextMatrix(fgPresu.Row, 3) = "R"
    'Verifica meses para distribucion.
    nTotalMesAnt = 0
    For N = 1 To nMesIni - 1
        nTotalMesAnt = nTotalMesAnt + CCur(IIf(fgPeriodo.TextMatrix(N, 3) = "", 0, fgPeriodo.TextMatrix(N, 3)))
    Next
    'Si es BALANCE
    If Trim(Right(cboTpo.Text, 1)) = "2" Then
        'nTotalPre = CCur(IIf(Trim(fgPresu.TextMatrix(fgPresu.Row, 9)) = "", 0, Trim(fgPresu.TextMatrix(fgPresu.Row, 9))))
        nTotalMes = Round(((nTotalPre - nTotalMesAnt) / (12 - nMesIni + 1)), 2)
    Else
        nTotalMes = Round(((nTotalPre - nTotalMesAnt) / (12 - nMesIni + 1)), 2)
    End If
    
    For N = nMesIni To fgPeriodo.Rows - 2
        fgPeriodo.TextMatrix(N, 3) = Format(nTotalMes, "#,##0.00")
    Next
    'OJO - Aqui me quede 02/06/2001
    nTotalMeses = 0
    For N = 1 To fgPeriodo.Rows - 2
        nTotalMeses = nTotalMeses + CCur(IIf(fgPeriodo.TextMatrix(N, 3) = "", 0, fgPeriodo.TextMatrix(N, 3)))
    Next
    fgPeriodo.TextMatrix(fgPeriodo.Rows - 1, 3) = Format(nTotalPre - nTotalMeses, "#,##0.00")
    'Suma TOTAL de MESES
    nTotalMeses = 0
    For N = 1 To fgPeriodo.Rows - 1
        nTotalMeses = nTotalMeses + CCur(IIf(fgPeriodo.TextMatrix(N, 3) = "", 0, fgPeriodo.TextMatrix(N, 3)))
    Next
    lblTotal.Caption = Format(nTotalMeses, "#,##0.00")
End If
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub fgPeriodo_DblClick()
Dim nTotalPre As Currency
nTotalPre = CCur(IIf(Trim(fgPresu.TextMatrix(fgPresu.Row, 5)) = "", 0, Trim(fgPresu.TextMatrix(fgPresu.Row, 5)))) + _
    CCur(IIf(Trim(fgPresu.TextMatrix(fgPresu.Row, 6)) = "", 0, Trim(fgPresu.TextMatrix(fgPresu.Row, 6)))) + _
    CCur(IIf(Trim(fgPresu.TextMatrix(fgPresu.Row, 7)) = "", 0, Trim(fgPresu.TextMatrix(fgPresu.Row, 7))))
If fgPeriodo.TextMatrix(fgPeriodo.Row, 0) <> "" And nTotalPre > 0 Then
    If fgPeriodo.Col >= 3 And fgPresu.TextMatrix(fgPresu.Row, 3) <> "G" Then
        EnfocaTexto txtObj2, 0, fgPeriodo
    End If
End If
End Sub
Private Sub fgPeriodo_GotFocus()
If txtObj2.Visible Then
   txtObj2.Visible = False
End If
End Sub
Private Sub fgPeriodo_KeyPress(KeyAscii As Integer)
If fgPeriodo.Col = 3 Then
    If KeyAscii = 13 Then EnfocaTexto txtObj2, IIf(KeyAscii = 13, 0, KeyAscii), fgPeriodo
End If
End Sub
Private Sub fgPeriodo_KeyUp(KeyCode As Integer, Shift As Integer)
If fgPeriodo.Col >= 4 Then
    KeyUp_Flex fgPeriodo, KeyCode, Shift
End If
End Sub
Private Sub fgPeriodo_RowColChange()
Dim nTotalPre As Currency
nTotalPre = CCur(IIf(Trim(fgPresu.TextMatrix(fgPresu.Row, 5)) = "", 0, Trim(fgPresu.TextMatrix(fgPresu.Row, 5)))) + _
    CCur(IIf(Trim(fgPresu.TextMatrix(fgPresu.Row, 6)) = "", 0, Trim(fgPresu.TextMatrix(fgPresu.Row, 6)))) + _
    CCur(IIf(Trim(fgPresu.TextMatrix(fgPresu.Row, 7)) = "", 0, Trim(fgPresu.TextMatrix(fgPresu.Row, 7))))
If fgPeriodo.TextMatrix(fgPeriodo.Row, 0) <> "" And nTotalPre > 0 Then
    If fgPeriodo.Col = 3 And fgPresu.TextMatrix(fgPresu.Row, 3) <> "G" Then
        fgPeriodo.FocusRect = flexFocusNone
        txtObj2.Left = fgPeriodo.Left + fgPeriodo.CellLeft - 15
        txtObj2.Top = fgPeriodo.Top + fgPeriodo.CellTop - 40
        txtObj2.Width = fgPeriodo.CellWidth
        txtObj2.Text = fgPeriodo
        txtObj2.Visible = True
    Else
        txtObj2.Visible = False
        fgPeriodo.FocusRect = flexFocusHeavy
    End If
End If
End Sub

Private Sub fgPresu_DblClick()
If fgPresu.TextMatrix(fgPresu.Row, 0) <> "" Then
    If IsUltNivel(Left(Trim(cboFecha.Text), 4), Right(cboPresu.Text, 4), fgPresu.TextMatrix(fgPresu.Row, 1)) And _
    fgPresu.TextMatrix(fgPresu.Row, 3) <> "G" And fgPresu.Col >= 4 And fgPresu.Col <= 7 Then
        EnfocaTexto txtObj, 0, fgPresu
    End If
End If
pRow = fgPresu.Row
End Sub

Private Sub fgPresu_GotFocus()
If txtObj.Visible Then
   txtObj.Visible = False
End If
End Sub
Private Sub fgPresu_KeyPress(KeyAscii As Integer)
If IsUltNivel(Left(Trim(cboFecha.Text), 4), Right(cboPresu.Text, 4), fgPresu.TextMatrix(fgPresu.Row, 1)) And _
fgPresu.TextMatrix(fgPresu.Row, 3) <> "G" And fgPresu.Col >= 4 And fgPresu.Col <= 7 Then
    If KeyAscii = 13 Then EnfocaTexto txtObj, IIf(KeyAscii = 13, 0, KeyAscii), fgPresu
End If
pRow = fgPresu.Row
End Sub
Private Sub fgPresu_KeyUp(KeyCode As Integer, Shift As Integer)
If IsUltNivel(Left(Trim(cboFecha.Text), 4), Right(cboPresu.Text, 4), fgPresu.TextMatrix(fgPresu.Row, 1)) And _
fgPresu.Col >= 4 And fgPresu.Col <= 7 Then
    KeyUp_Flex fgPresu, KeyCode, Shift
End If
pRow = fgPresu.Row
End Sub

Private Sub fgPresu_RowColChange()
If fgPresu.TextMatrix(fgPresu.Row, 0) <> "" Then
    If fgPresu.Row <> pRow And fgPresu.TextMatrix(pRow, 3) = "M" Then
        fgPresu.Row = pRow
        fgPresu.SetFocus
        Exit Sub
    End If
    'Meses
    lblTotal.Caption = ""
    Call CargaPeriodo
    If IsUltNivel(Left(Trim(cboFecha.Text), 4), Right(cboPresu.Text, 4), fgPresu.TextMatrix(fgPresu.Row, 1)) And _
       fgPresu.Col >= 4 And fgPresu.Col <= 7 Then
        txtObj2.Visible = False
        fgPresu.FocusRect = flexFocusNone
        If fgPresu.TextMatrix(fgPresu.Row, 3) = "G" Then
            txtObj.Visible = False
            cmdModificar.Enabled = True
            cmdMov.Enabled = False
            cmdGrabar.Enabled = False
            cmdCancelar.Enabled = False
            pRow = fgPresu.Row
            Exit Sub
        End If
        cmdModificar.Enabled = False
        cmdMov.Enabled = True
        cmdGrabar.Enabled = True
        If fgPresu.TextMatrix(fgPresu.Row, 3) = "M" Then
            cmdCancelar.Enabled = True
        Else
            cmdCancelar.Enabled = False
        End If
        txtObj.Text = fgPresu
        txtObj.Left = fgPresu.Left + fgPresu.CellLeft - 15
        txtObj.Top = fgPresu.Top + fgPresu.CellTop - 15
        txtObj.Width = fgPresu.CellWidth
        txtObj.Visible = True
    Else
        cmdModificar.Enabled = False
        cmdMov.Enabled = False
        cmdGrabar.Enabled = False
        cmdCancelar.Enabled = False
        txtObj.Visible = False
        txtObj2.Visible = False
        fgPresu.FocusRect = flexFocusHeavy
    End If
End If
pRow = fgPresu.Row
End Sub

Private Sub fgPresu_Scroll()
txtObj.Visible = False
End Sub

Private Sub Form_Load()
    Dim tmpSql As String
    Dim x As Integer, N As Integer
    Dim oPP As DPresupuesto
    Set oPP = New DPresupuesto
    Dim oCon As DConstantes
    Set oCon = New DConstantes
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim clsDGnral As DLogGeneral
    Set clsDGnral = New DLogGeneral
    CentraForm Me
    pbIni = False
    'Limpiar
    'Carga los Años
    Set rs = clsDGnral.CargaPeriodo
    Call CargaCombo(rs, cboFecha)
    'Carga las Monedas
    Set rs = clsDGnral.CargaConstante(gMoneda, False)
    Call CargaCombo(rs, cboMoneda)    'Carga el nombre de los Presupuestos
    Call CargaCombo(oPP.GetPresupuesto(True), cboPresu, , 1, 0)
    pbIni = True
    Set rs = oCon.GetConstante(gPPPresupuestoTpo)
    CargaCombo rs, cboTpo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CierraConexion
End Sub

Private Sub txtMes_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub txtObj_GotFocus()
    txtObj.Width = fgPresu.CellWidth '- cmdExa.Width + 60
End Sub

Private Sub txtObj_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Or KeyCode = 38 Then
       txtObj_KeyPress 13
       SendKeys IIf(KeyCode = 38, "{Up}", "{Down}"), True
    End If
End Sub

Private Sub txtObj_KeyPress(KeyAscii As Integer)
    Dim N As Integer
    Dim nTot As Currency
    Dim nTotAnt As Currency
    KeyAscii = NumerosDecimales(txtObj, KeyAscii, 12)
    If KeyAscii = 13 Then
        txtObj2.Visible = False
        Call CuadraFil
        fgPresu.Enabled = True
        txtObj.Visible = False
        fgPresu.TextMatrix(fgPresu.Row, fgPresu.Col) = Format(txtObj.Text, "#,##0.00")
        'Total - Variación
        nTot = 0
        For N = 5 To 7
            nTot = nTot + CCur(IIf(fgPresu.TextMatrix(fgPresu.Row, N) = "", 0, fgPresu.TextMatrix(fgPresu.Row, N)))
        Next
        nTotAnt = CCur(IIf(fgPresu.TextMatrix(fgPresu.Row, 4) = "", 0, fgPresu.TextMatrix(fgPresu.Row, 4)))
        fgPresu.TextMatrix(fgPresu.Row, 8) = Format(nTot, "#,##0.00")
        If nTot <> 0 And nTotAnt <> 0 Then
            fgPresu.TextMatrix(fgPresu.Row, 9) = Format(nTot - nTotAnt, "#,##0.00")
            fgPresu.TextMatrix(fgPresu.Row, 10) = Format(Round((((nTot / nTotAnt) - 1) * 100), 2), "#,##0.00")
        ElseIf nTot <> 0 Or nTotAnt <> 0 Then
            fgPresu.TextMatrix(fgPresu.Row, 9) = Format(nTot - nTotAnt, "#,##0.00")
            If nTot = 0 Then
                fgPresu.TextMatrix(fgPresu.Row, 10) = Format(100 * -1, "#,##0.00")
            Else
                fgPresu.TextMatrix(fgPresu.Row, 10) = Format(100, "#,##0.00")
            End If
        Else
            fgPresu.TextMatrix(fgPresu.Row, 9) = ""
            fgPresu.TextMatrix(fgPresu.Row, 10) = ""
        End If
        
        If CCur(IIf(txtObj.Text = "", 0, txtObj.Text)) = 0 Then
            For N = 1 To fgPeriodo.Rows - 1
                fgPeriodo.TextMatrix(N, 3) = ""
            Next
        End If
        fgPresu.SetFocus
    End If
End Sub

Private Sub txtObj_Validate(Cancel As Boolean)
    fgPresu.TextMatrix(fgPresu.Row, fgPresu.Col) = Format(txtObj.Text, "#,##0.00")
End Sub

Private Sub txtObj2_GotFocus()
txtObj2.Width = fgPeriodo.CellWidth '- cmdExa.Width + 60
End Sub

Private Sub txtObj2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Or KeyCode = 38 Then
       txtObj2_KeyPress 13
       SendKeys IIf(KeyCode = 38, "{Up}", "{Down}"), True
    End If
End Sub
Private Sub txtObj2_KeyPress(KeyAscii As Integer)
    Dim N As Integer
    Dim vMonto As Currency
    Dim nTotalPre As Currency
    KeyAscii = NumerosDecimales(txtObj2, KeyAscii, 12, 2)
    If KeyAscii = 13 Then
        txtObj2.Visible = False
        fgPeriodo.TextMatrix(fgPeriodo.Row, fgPeriodo.Col) = Format(txtObj2.Text, "#,##0.00")
        fgPeriodo.SetFocus
        lblTotal.Caption = ""
        For N = 1 To fgPeriodo.Rows - 1
            vMonto = vMonto + CCur(IIf(fgPeriodo.TextMatrix(N, 3) = "", 0, fgPeriodo.TextMatrix(N, 3)))
        Next
        'BALANCE
        If Trim(Right(cboTpo.Text, 1)) = "2" Then
            nTotalPre = CCur(IIf(Trim(fgPresu.TextMatrix(fgPresu.Row, 9)) = "", 0, Trim(fgPresu.TextMatrix(fgPresu.Row, 9))))
        Else
            nTotalPre = CCur(IIf(Trim(fgPresu.TextMatrix(fgPresu.Row, 5)) = "", 0, Trim(fgPresu.TextMatrix(fgPresu.Row, 5)))) + _
                CCur(IIf(Trim(fgPresu.TextMatrix(fgPresu.Row, 6)) = "", 0, Trim(fgPresu.TextMatrix(fgPresu.Row, 6)))) + _
                CCur(IIf(Trim(fgPresu.TextMatrix(fgPresu.Row, 7)) = "", 0, Trim(fgPresu.TextMatrix(fgPresu.Row, 7))))
        End If
        If vMonto > nTotalPre Then
            MsgBox "Cantidades exceden el monto total", vbInformation, " Aviso "
            vMonto = vMonto - CCur(txtObj2.Text)
            fgPeriodo.TextMatrix(fgPeriodo.Row, 3) = ""
        End If
        lblTotal.Caption = Format(vMonto, "#,##0.00")
    End If
End Sub

Private Sub txtObj2_Validate(Cancel As Boolean)
    Dim nTotalPre As Currency
    Dim vMonto As Currency
    Dim N As Integer
    fgPeriodo.TextMatrix(fgPeriodo.Row, fgPeriodo.Col) = Format(txtObj2.Text, "#,##0.00")
    lblTotal.Caption = ""
    For N = 1 To fgPeriodo.Rows - 1
        vMonto = vMonto + CCur(IIf(fgPeriodo.TextMatrix(N, 3) = "", 0, fgPeriodo.TextMatrix(N, 3)))
    Next
    nTotalPre = CCur(IIf(Trim(fgPresu.TextMatrix(fgPresu.Row, 5)) = "", 0, Trim(fgPresu.TextMatrix(fgPresu.Row, 5)))) + _
        CCur(IIf(Trim(fgPresu.TextMatrix(fgPresu.Row, 6)) = "", 0, Trim(fgPresu.TextMatrix(fgPresu.Row, 6)))) + _
        CCur(IIf(Trim(fgPresu.TextMatrix(fgPresu.Row, 7)) = "", 0, Trim(fgPresu.TextMatrix(fgPresu.Row, 7))))
    
    If vMonto > nTotalPre Then
        vMonto = vMonto - CCur(txtObj2.Text)
        fgPeriodo.TextMatrix(fgPeriodo.Row, 3) = ""
    End If
    lblTotal.Caption = Format(vMonto, "#,##0.00")
End Sub

Private Sub Limpiar()
txtObj.Visible = False
txtObj2.Visible = False
'If Right(Trim(cboTpo.Text), 1) = "2" Then
    Call MSHFlex(fgPresu, 11, "Item-Código-Descripción-Opc-Año " & Val(Trim(cboFecha.Text)) - 1 & " -Monto Presup-Cred. 01-Cred. 02-Total-Var.Monto-Var. %", "400-1100-2200-400-1100-1100-1000-1000-1000-1000-800", "R-L-L-C-R-R-R-R-R-R-R")
'Else
'    Call MSHFlex(fgPresu, 8, "Item-Código-Descripción-Opc-Monto Inicial-Monto Presup-Cred. 01-Cred. 02", "450-1100-2600-400-0-1100-1000-1000", "R-L-L-C-R-R-R-R")
'End If
Call MSHFlex(fgPeriodo, 4, "Item-Código-Mes-Monto", "400-0-900-1300", "R-L-L-R")
End Sub

Private Sub BloqueoCmd()
cmdGrabar.Enabled = False
cmdMov.Enabled = False
cmdModificar.Enabled = False
cmdCancelar.Enabled = False
End Sub

Private Sub CuadraFil()
Dim x As Integer
If fgPresu.TextMatrix(fgPresu.Row, 3) = "R" Then
    For x = 1 To 12
        fgPeriodo.TextMatrix(x, 3) = ""
    Next
End If
End Sub

Private Sub CargaPeriodo()
    Dim tmpReg As ADODB.Recordset
    Set tmpReg = New ADODB.Recordset
    Dim oPP As DPresupuesto
    Set oPP = New DPresupuesto
    Dim tmpSql As String
    Dim x As Integer, N As Integer
    Dim vMonto As Currency
    fgPeriodo.Redraw = False
    Call MSHFlex(fgPeriodo, 4, "Item-Código-Mes-Monto", "400-0-900-1300", "R-L-L-R")
    
    Set tmpReg = oPP.GetPresupRubroMes(cboFecha.Text, Right(Trim(cboPresu.Text), 4), Right(Trim(cboMoneda.Text), 1), fgPresu.TextMatrix(fgPresu.Row, 1))

    If (tmpReg.BOF Or tmpReg.EOF) Then
    Else
        With tmpReg
            Do While Not .EOF
                x = x + 1
                AdicionaRow fgPeriodo, x
                fgPeriodo.Row = fgPeriodo.Rows - 1
                fgPeriodo.TextMatrix(x, 0) = x
                fgPeriodo.TextMatrix(x, 1) = !nConsValor
                fgPeriodo.TextMatrix(x, 2) = !cNomtab
                fgPeriodo.TextMatrix(x, 3) = IIf(IsNull(!nPresuRubMesMonIni), "", Format(!nPresuRubMesMonIni, "#,##0.00"))
                .MoveNext
            Loop
        End With
        lblTotal.Caption = ""
        For N = 1 To fgPeriodo.Rows - 1
            vMonto = vMonto + CCur(IIf(fgPeriodo.TextMatrix(N, 3) = "", 0, fgPeriodo.TextMatrix(N, 3)))
        Next
    End If
    tmpReg.Close
    Set tmpReg = Nothing
    fgPeriodo.Redraw = True
    lblTotal.Caption = Format(vMonto, "#,##0.00")
End Sub

Private Sub CargaPresu()
    Dim tmpReg As ADODB.Recordset
    Set tmpReg = New ADODB.Recordset
    Dim oPP As DPresupuesto
    Set oPP = New DPresupuesto
    Dim tmpSql As String
    Dim x As Integer, N As Integer
    fgPresu.Redraw = False
    
    Set tmpReg = oPP.GetPresupRubro(cboFecha.Text, Right(cboPresu.Text, 4))
    If Not (tmpReg.BOF Or tmpReg.EOF) Then
        With tmpReg
            Do While Not .EOF
                x = x + 1
                'fgPresu.Rows = X + 1
                AdicionaRow fgPresu, x
                fgPresu.Row = fgPresu.Rows - 1
                fgPresu.TextMatrix(x, 0) = x
                fgPresu.TextMatrix(x, 1) = !cPresuRubCod
                fgPresu.TextMatrix(x, 2) = !cPresuRubDescripcion
                fgPresu.TextMatrix(x, 3) = IIf(!nTotal + !nMonIni = 0, "", "G")
                If !nMonIni <> 0 Or !nMonto <> 0 Then
                    fgPresu.TextMatrix(x, 4) = Format(!nMonIni, "#,##0.00")
                    fgPresu.TextMatrix(x, 5) = Format(!nMonto, "#,##0.00")
                Else
                    fgPresu.TextMatrix(x, 4) = IIf(!nMonIni = 0, "", Format(!nMonIni, "#,##0.00"))
                    fgPresu.TextMatrix(x, 5) = IIf(!nMonto = 0, "", Format(!nMonto, "#,##0.00"))
                End If
                fgPresu.TextMatrix(x, 6) = IIf(!nMonCre1 = 0, "", Format(!nMonCre1, "#,##0.00"))
                fgPresu.TextMatrix(x, 7) = IIf(!nMonCre2 = 0, "", Format(!nMonCre2, "#,##0.00"))
                If !nMonIni <> 0 Or !nMonto <> 0 Then
                    fgPresu.TextMatrix(x, 8) = Format(!nTotal, "#,##0.00")
                Else
                    fgPresu.TextMatrix(x, 8) = IIf(!nTotal = 0, "", Format(!nTotal, "#,##0.00"))
                End If
                If !nTotal <> 0 And !nMonIni <> 0 Then
                    fgPresu.TextMatrix(x, 9) = Format(!nTotal - !nMonIni, "#,##0.00")
                    fgPresu.TextMatrix(x, 10) = Format(Round((((!nTotal / !nMonIni) - 1) * 100), 2), "#,##0.00")
                ElseIf !nTotal <> 0 Or !nMonIni <> 0 Then
                    fgPresu.TextMatrix(x, 9) = Format(!nTotal - !nMonIni, "#,##0.00")
                    If !nTotal = 0 Then
                        fgPresu.TextMatrix(x, 10) = Format(100 * -1, "#,##0.00")
                    Else
                        fgPresu.TextMatrix(x, 10) = Format(100, "#,##0.00")
                    End If
                Else
                    fgPresu.TextMatrix(x, 9) = ""
                    fgPresu.TextMatrix(x, 10) = ""
                End If
                .MoveNext
            Loop
        End With
    End If
    tmpReg.Close
    Set tmpReg = Nothing
    fgPresu.Redraw = True
    pRow = fgPresu.Row
End Sub

Private Sub Cabecera(ByVal cTipo As String, ByVal nPage As Integer)
If nPage > 1 Then vRTFImp = vRTFImp & oImpresora.gPrnSaltoPagina
vRTFImp = vRTFImp & "  CMAC - TRUJILLO" & Space(140) & Format(gdFecSis & " " & Time, gsFormatoFechaHoraView) & oImpresora.gPrnSaltoLinea
vRTFImp = vRTFImp & ImpreFormat(UCase(gsNomAge), 25) & Space(135) & " Página :" & ImpreFormat(nPage, 5, 0) & oImpresora.gPrnSaltoLinea
If cTipo = "1" Then
    vRTFImp = vRTFImp & Space(55) & "LISTADO DEL AÑO " & Left(Trim(cboFecha), 4) & " DE : " & UCase(Left(cboPresu, 25)) & " EN " & Left(Trim(cboMoneda), 12) & oImpresora.gPrnSaltoLinea
End If
vRTFImp = vRTFImp & Space(2) & String(175, "-") & oImpresora.gPrnSaltoLinea
vRTFImp = vRTFImp & Space(2) & "     CODIGO            DESCRIPCION             ENERO      FEBRE      MARZO      ABRIL      MAYO       JUNIO      JULIO      AGOST      SETIE      OCTUB      NOVIE      DICIE" & oImpresora.gPrnSaltoLinea
vRTFImp = vRTFImp & Space(2) & String(175, "-") & oImpresora.gPrnSaltoLinea
End Sub

