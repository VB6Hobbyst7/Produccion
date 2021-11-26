VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPlaEjecu 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11865
   Icon            =   "frmPlaEjecu.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkMovContables 
      Appearance      =   0  'Flat
      Caption         =   "Comp. Mov Cont."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5250
      TabIndex        =   23
      Top             =   5955
      Width           =   1950
   End
   Begin VB.CommandButton cmdEjecutar 
      Cancel          =   -1  'True
      Caption         =   "&Ejecutar"
      Height          =   390
      Left            =   9345
      TabIndex        =   7
      Top             =   5850
      Width           =   1170
   End
   Begin VB.CheckBox chkMensuales 
      Appearance      =   0  'Flat
      Caption         =   "Presup. &Mensual"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5250
      TabIndex        =   22
      Top             =   5670
      Width           =   1950
   End
   Begin VB.CheckBox chkCierreAño 
      Appearance      =   0  'Flat
      Caption         =   "&Cierre Año"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2775
      TabIndex        =   21
      Top             =   6255
      Value           =   1  'Checked
      Width           =   1230
   End
   Begin VB.ComboBox cboMonedaP 
      Height          =   315
      ItemData        =   "frmPlaEjecu.frx":08CA
      Left            =   9675
      List            =   "frmPlaEjecu.frx":08CC
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   83
      Width           =   2115
   End
   Begin VB.CheckBox chkBala 
      Appearance      =   0  'Flat
      Caption         =   "Con Saldos de Balance"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2775
      TabIndex        =   18
      Top             =   5925
      Width           =   2355
   End
   Begin VB.ComboBox cboFecha 
      Height          =   315
      Left            =   660
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   83
      Width           =   1000
   End
   Begin VB.ComboBox cboPresu 
      Height          =   315
      Left            =   2580
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   83
      Width           =   3435
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   10575
      TabIndex        =   9
      Top             =   5850
      Width           =   1170
   End
   Begin VB.ComboBox cboMoneda 
      Enabled         =   0   'False
      Height          =   315
      Left            =   9420
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   660
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.ComboBox cboTpo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6825
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   83
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Impresión en hoja de Cálculo "
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   75
      TabIndex        =   1
      Top             =   5565
      Width           =   2625
      Begin VB.OptionButton optImp 
         Appearance      =   0  'Flat
         Caption         =   "Semestral"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   105
         TabIndex        =   5
         Top             =   645
         Width           =   1020
      End
      Begin VB.OptionButton optImp 
         Appearance      =   0  'Flat
         Caption         =   "Trimestral"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   105
         TabIndex        =   4
         Top             =   420
         Width           =   1020
      End
      Begin VB.OptionButton optImp 
         Appearance      =   0  'Flat
         Caption         =   "Mensual"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   105
         TabIndex        =   3
         Top             =   195
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.CommandButton cmdImprimir 
         Height          =   390
         Left            =   1485
         Picture         =   "frmPlaEjecu.frx":08CE
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   390
         Width           =   1050
      End
   End
   Begin VB.CheckBox chkProy 
      Appearance      =   0  'Flat
      Caption         =   "Con gastos comprometidos."
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2775
      TabIndex        =   0
      Top             =   5640
      Width           =   2355
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgPresu 
      Height          =   4995
      Left            =   90
      TabIndex        =   12
      Top             =   510
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   8811
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
   Begin RichTextLib.RichTextBox rtfImp 
      Height          =   315
      Left            =   11250
      TabIndex        =   13
      Top             =   6150
      Visible         =   0   'False
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   556
      _Version        =   393217
      TextRTF         =   $"frmPlaEjecu.frx":15D4
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
      ForeColor       =   &H00000040&
      Height          =   270
      Index           =   2
      Left            =   8805
      TabIndex        =   20
      Top             =   105
      Width           =   855
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
      ForeColor       =   &H00000040&
      Height          =   210
      Index           =   0
      Left            =   150
      TabIndex        =   17
      Top             =   135
      Width           =   465
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
      ForeColor       =   &H00000040&
      Height          =   210
      Index           =   1
      Left            =   1725
      TabIndex        =   16
      Top             =   135
      Width           =   810
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
      Left            =   9420
      TabIndex        =   15
      Top             =   525
      Visible         =   0   'False
      Width           =   825
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
      ForeColor       =   &H00000040&
      Height          =   270
      Index           =   4
      Left            =   6255
      TabIndex        =   14
      Top             =   105
      Width           =   510
   End
End
Attribute VB_Name = "frmPlaEjecu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vRTFImp As String
Dim pbIni As Boolean

Private Sub cboFecha_Click()
If pbIni Then
    fgPresu.FixedRows = 1
    Call Limpiar
    Call CargaPresu(False, Switch(optImp(0).value = True, 1, optImp(1).value = True, 2, optImp(2).value = True, 3))
    If fgPresu.Rows > 2 Then fgPresu.FixedRows = 2
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
    fgPresu.FixedRows = 1
    Call Limpiar
    Call CargaPresu(False, Switch(optImp(0).value = True, 1, optImp(1).value = True, 2, optImp(2).value = True, 3))
    If fgPresu.Rows > 2 Then fgPresu.FixedRows = 2
End Sub

Private Sub cmdEjecutar_Click()
    MousePointer = 11
    fgPresu.FixedRows = 1
    If Not optImp(0).value Then
        MsgBox "Vista No Permitida, Solo se puede Imprimir", vbInformation, "Aviso"
        Exit Sub
    End If
    Call Limpiar
    Call CargaPresu(True, Switch(optImp(0).value = True, 1, optImp(1).value = True, 2, optImp(2).value = True, 3))
    If fgPresu.Rows > 2 Then fgPresu.FixedRows = 2
    MousePointer = 0
End Sub

Private Sub CmdImprimir_Click()
    Dim sArchivo As String
    MousePointer = 11

    sArchivo = "PresuEjec" & cboFecha & ".XLS"
    
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

    Dim tmpReg As New ADODB.Recordset
    Dim tmpSql As String
    Dim nFil As Integer, nCol As Integer, nItem As Integer
    Dim TotMov As Currency
    Dim nMonBal As Currency, nMonMov As Currency
    Dim nMonSdo As Currency
    
    Dim oCon As DConecta
    Set oCon = New DConecta
    oCon.AbreConexion
    
    If Not ExcelBegin(App.path & "\SPOOLER\" & psArchivo, xlAplicacion, xlLibro, False) Then
        Exit Sub
    End If
    lbExcel = True
    ExcelAddHoja Left(Me.cboPresu, 20), xlLibro, xlHoja1
    xlHoja1.PageSetup.Zoom = 70
    xlHoja1.PageSetup.Orientation = xlLandscape
    
    xlAplicacion.Range("A1:BJ100").Font.Size = 7
    
    xlHoja1.Range("A1").ColumnWidth = 5
    xlHoja1.Range("B1").ColumnWidth = 10
    xlHoja1.Range("C1").ColumnWidth = 35
    xlHoja1.Range("D1..J1").ColumnWidth = 11
    xlHoja1.Range("K1..BJ1").ColumnWidth = 10
    
    xlHoja1.Cells(1, 1) = gsNomCmac
    xlAplicacion.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 1)).Font.Bold = True
    If Right(Trim(cboTpo.Text), 1) = "2" Then
        'SOLO BALANCES
        If pnTipo = 1 Then
            xlHoja1.Cells(1, 45) = "Fecha :" & Format(gdFecSis, "dd mmmm yyyy")
            xlAplicacion.Range(xlHoja1.Cells(1, 45), xlHoja1.Cells(1, 45)).Font.Bold = True
        ElseIf pnTipo = 2 Then
            xlHoja1.Cells(1, 21) = "Fecha :" & Format(gdFecSis, "dd mmmm yyyy")
            xlAplicacion.Range(xlHoja1.Cells(1, 21), xlHoja1.Cells(1, 21)).Font.Bold = True
        Else
            xlHoja1.Cells(1, 15) = "Fecha :" & Format(gdFecSis, "dd mmmm yyyy")
            xlAplicacion.Range(xlHoja1.Cells(1, 15), xlHoja1.Cells(1, 15)).Font.Bold = True
        End If
    Else
        If pnTipo = 1 Then
            xlHoja1.Cells(1, 60) = "Fecha :" & Format(gdFecSis, "dd mmmm yyyy")
            xlAplicacion.Range(xlHoja1.Cells(1, 60), xlHoja1.Cells(1, 60)).Font.Bold = True
        ElseIf pnTipo = 2 Then
            xlHoja1.Cells(1, 28) = "Fecha :" & Format(gdFecSis, "dd mmmm yyyy")
            xlAplicacion.Range(xlHoja1.Cells(1, 28), xlHoja1.Cells(1, 28)).Font.Bold = True
        Else
            xlHoja1.Cells(1, 20) = "Fecha :" & Format(gdFecSis, "dd mmmm yyyy")
            xlAplicacion.Range(xlHoja1.Cells(1, 20), xlHoja1.Cells(1, 20)).Font.Bold = True
        End If
    End If
    xlHoja1.Cells(2, 1) = "Area de Planeamiento"
    xlAplicacion.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 1)).Font.Bold = True
    If Right(Trim(cboTpo.Text), 1) = "2" Then
        'SOLO BALANCES
        If pnTipo = 1 Then
            xlHoja1.Cells(3, 20) = Trim(Mid(cboPresu.Text, 1, Len(cboPresu.Text) - 4))
            xlAplicacion.Range(xlHoja1.Cells(3, 20), xlHoja1.Cells(3, 20)).Font.Bold = True
            
            xlHoja1.Cells(4, 21) = " AÑO : " & Trim(cboFecha.Text)
            xlAplicacion.Range(xlHoja1.Cells(4, 21), xlHoja1.Cells(4, 21)).Font.Bold = True
        ElseIf pnTipo = 2 Then
            xlHoja1.Cells(3, 9) = Trim(Mid(cboPresu.Text, 1, Len(cboPresu.Text) - 4))
            xlAplicacion.Range(xlHoja1.Cells(3, 9), xlHoja1.Cells(3, 9)).Font.Bold = True
            
            xlHoja1.Cells(4, 10) = " AÑO : " & Trim(cboFecha.Text)
            xlAplicacion.Range(xlHoja1.Cells(4, 10), xlHoja1.Cells(4, 10)).Font.Bold = True
        Else
            xlHoja1.Cells(3, 7) = Trim(Mid(cboPresu.Text, 1, Len(cboPresu.Text) - 4))
            xlAplicacion.Range(xlHoja1.Cells(3, 7), xlHoja1.Cells(3, 7)).Font.Bold = True
            
            xlHoja1.Cells(4, 8) = " AÑO : " & Trim(cboFecha.Text)
            xlAplicacion.Range(xlHoja1.Cells(4, 8), xlHoja1.Cells(4, 8)).Font.Bold = True
        End If
    Else
        If pnTipo = 1 Then
            xlHoja1.Cells(3, 26) = Trim(Mid(cboPresu.Text, 1, Len(cboPresu.Text) - 4))
            xlAplicacion.Range(xlHoja1.Cells(3, 26), xlHoja1.Cells(3, 26)).Font.Bold = True
            
            xlHoja1.Cells(4, 27) = " AÑO : " & Trim(cboFecha.Text)
            xlAplicacion.Range(xlHoja1.Cells(4, 27), xlHoja1.Cells(4, 27)).Font.Bold = True
        ElseIf pnTipo = 2 Then
            xlHoja1.Cells(3, 11) = Trim(Mid(cboPresu.Text, 1, Len(cboPresu.Text) - 4))
            xlAplicacion.Range(xlHoja1.Cells(3, 11), xlHoja1.Cells(3, 11)).Font.Bold = True
            
            xlHoja1.Cells(4, 12) = " AÑO : " & Trim(cboFecha.Text)
            xlAplicacion.Range(xlHoja1.Cells(4, 12), xlHoja1.Cells(4, 12)).Font.Bold = True
        Else
            xlHoja1.Cells(3, 8) = Trim(Mid(cboPresu.Text, 1, Len(cboPresu.Text) - 4))
            xlAplicacion.Range(xlHoja1.Cells(3, 8), xlHoja1.Cells(3, 8)).Font.Bold = True
            
            xlHoja1.Cells(4, 9) = " AÑO : " & Trim(cboFecha.Text)
            xlAplicacion.Range(xlHoja1.Cells(4, 9), xlHoja1.Cells(4, 9)).Font.Bold = True
        End If
    End If
    xlAplicacion.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(7, 62)).HorizontalAlignment = xlHAlignCenter
    xlAplicacion.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(7, 62)).Font.Bold = True
    nCol = 1
    xlHoja1.Cells(7, 1) = "ITEM": xlHoja1.Cells(7, 2) = "CODIGO": xlHoja1.Cells(7, 3) = "DESCRIPCION": xlHoja1.Cells(7, 4) = "AÑO " & Val(Trim(cboFecha.Text)) - 1: xlHoja1.Cells(7, 5) = "PRESUPUESTO": xlHoja1.Cells(7, 6) = "CRED.1": xlHoja1.Cells(7, 7) = "CRED.2": xlHoja1.Cells(7, 8) = "TOTAL": xlHoja1.Cells(7, 9) = "Var.Monto": xlHoja1.Cells(7, 10) = "Var. %"
    
    CuadroExcel xlHoja1, 1, 6, 10, 7
    If Right(Trim(cboTpo.Text), 1) = "2" Then
        'SOLO BALANCES
        If pnTipo = 1 Then
            CuadroExcel xlHoja1, 11, 6, 13, 6, True: CuadroExcel xlHoja1, 13, 6, 16, 6, True
            CuadroExcel xlHoja1, 16, 6, 19, 6, True: CuadroExcel xlHoja1, 19, 6, 22, 6, True
            CuadroExcel xlHoja1, 22, 6, 25, 6, True: CuadroExcel xlHoja1, 25, 6, 28, 6, True
            CuadroExcel xlHoja1, 28, 6, 31, 6, True: CuadroExcel xlHoja1, 31, 6, 34, 6, True
            CuadroExcel xlHoja1, 34, 6, 37, 6, True: CuadroExcel xlHoja1, 37, 6, 40, 6, True
            CuadroExcel xlHoja1, 40, 6, 43, 6, True: CuadroExcel xlHoja1, 43, 6, 46, 6, True
            CuadroExcel xlHoja1, 46, 6, 47, 6, True
            CuadroExcel xlHoja1, 11, 7, 47, 7
            
            xlHoja1.Cells(6, 12) = "ENERO": xlHoja1.Cells(6, 15) = "FEBRERO": xlHoja1.Cells(6, 18) = "MARZO": xlHoja1.Cells(6, 21) = "ABRIL": xlHoja1.Cells(6, 24) = "MAYO": xlHoja1.Cells(6, 27) = "JUNIO": xlHoja1.Cells(6, 30) = "JULIO": xlHoja1.Cells(6, 33) = "AGOSTO": xlHoja1.Cells(6, 36) = "SETIEMBRE": xlHoja1.Cells(6, 39) = "OCTUBRE": xlHoja1.Cells(6, 42) = "NOVIEMBRE": xlHoja1.Cells(6, 45) = "DICIEMBRE": xlHoja1.Cells(6, 47) = "TOTALES"
            xlHoja1.Cells(7, 11) = "PRESUP. MES": xlHoja1.Cells(7, 12) = "MOVIM. MES": xlHoja1.Cells(7, 13) = "SALDO ENE": xlHoja1.Cells(7, 14) = "PRESUP. MES": xlHoja1.Cells(7, 15) = "MOVIM. MES": xlHoja1.Cells(7, 16) = "SALDO FEB"
            xlHoja1.Cells(7, 17) = "PRESUP. MES": xlHoja1.Cells(7, 18) = "MOVIM. MES": xlHoja1.Cells(7, 19) = "SALDO MAR": xlHoja1.Cells(7, 20) = "PRESUP. MES": xlHoja1.Cells(7, 21) = "MOVIM. MES": xlHoja1.Cells(7, 22) = "SALDO ABR"
            xlHoja1.Cells(7, 23) = "PRESUP. MES": xlHoja1.Cells(7, 24) = "MOVIM. MES": xlHoja1.Cells(7, 25) = "SALDO MAY": xlHoja1.Cells(7, 26) = "PRESUP. MES": xlHoja1.Cells(7, 27) = "MOVIM. MES": xlHoja1.Cells(7, 28) = "SALDO JUN"
            xlHoja1.Cells(7, 29) = "PRESUP. MES": xlHoja1.Cells(7, 30) = "MOVIM. MES": xlHoja1.Cells(7, 31) = "SALDO JUL": xlHoja1.Cells(7, 32) = "PRESUP. MES": xlHoja1.Cells(7, 33) = "MOVIM. MES": xlHoja1.Cells(7, 34) = "SALDO AGO"
            xlHoja1.Cells(7, 35) = "PRESUP. MES": xlHoja1.Cells(7, 36) = "MOVIM. MES": xlHoja1.Cells(7, 37) = "SALDO SET": xlHoja1.Cells(7, 38) = "PRESUP. MES": xlHoja1.Cells(7, 39) = "MOVIM. MES": xlHoja1.Cells(7, 40) = "SALDO OCT"
            xlHoja1.Cells(7, 41) = "PRESUP. MES": xlHoja1.Cells(7, 42) = "MOVIM. MES": xlHoja1.Cells(7, 43) = "SALDO NOV": xlHoja1.Cells(7, 44) = "PRESUP. MES": xlHoja1.Cells(7, 45) = "MOVIM. MES": xlHoja1.Cells(7, 46) = "SALDO DIC"
            xlHoja1.Cells(7, 47) = "AVANCE %"
        ElseIf pnTipo = 2 Then
            CuadroExcel xlHoja1, 11, 6, 13, 6, True: CuadroExcel xlHoja1, 13, 6, 16, 6, True
            CuadroExcel xlHoja1, 16, 6, 19, 6, True: CuadroExcel xlHoja1, 19, 6, 22, 6, True
            CuadroExcel xlHoja1, 22, 6, 23, 6, True
            CuadroExcel xlHoja1, 11, 7, 23, 7
            
            xlHoja1.Cells(6, 12) = "PRIMERO": xlHoja1.Cells(6, 15) = "SEGUNDO": xlHoja1.Cells(6, 18) = "TERCERO": xlHoja1.Cells(6, 21) = "CUARTO": xlHoja1.Cells(6, 23) = "TOTALES"
            xlHoja1.Cells(7, 11) = "PRESUP. MES": xlHoja1.Cells(7, 12) = "MOVIM. MES": xlHoja1.Cells(7, 13) = "SALDO I": xlHoja1.Cells(7, 14) = "PRESUP. MES": xlHoja1.Cells(7, 15) = "MOVIM. MES": xlHoja1.Cells(7, 16) = "SALDO II"
            xlHoja1.Cells(7, 17) = "PRESUP. MES": xlHoja1.Cells(7, 18) = "MOVIM. MES": xlHoja1.Cells(7, 19) = "SALDO III": xlHoja1.Cells(7, 20) = "PRESUP. MES": xlHoja1.Cells(7, 21) = "MOVIM. MES": xlHoja1.Cells(7, 22) = "SALDO IV"
            xlHoja1.Cells(7, 23) = "AVANCE %"
        Else
            CuadroExcel xlHoja1, 11, 6, 13, 6, True: CuadroExcel xlHoja1, 13, 6, 16, 6, True
            CuadroExcel xlHoja1, 16, 6, 17, 6, True
            CuadroExcel xlHoja1, 11, 7, 17, 7
            
            xlHoja1.Cells(6, 12) = "PRIMERO": xlHoja1.Cells(6, 15) = "SEGUNDO": xlHoja1.Cells(6, 17) = "TOTALES"
            xlHoja1.Cells(7, 11) = "PRESUP. MES": xlHoja1.Cells(7, 12) = "MOVIM. MES": xlHoja1.Cells(7, 13) = "SALDO I": xlHoja1.Cells(7, 14) = "PRESUP. MES": xlHoja1.Cells(7, 15) = "MOVIM. MES": xlHoja1.Cells(7, 16) = "SALDO II"
            xlHoja1.Cells(7, 17) = "AVANCE %"
        End If
    Else
        If pnTipo = 1 Then
            'CuadroExcel xlHoja1, 9, 6, 56, 6, True
            CuadroExcel xlHoja1, 11, 6, 14, 6, True: CuadroExcel xlHoja1, 14, 6, 18, 6, True
            CuadroExcel xlHoja1, 18, 6, 22, 6, True: CuadroExcel xlHoja1, 22, 6, 26, 6, True
            CuadroExcel xlHoja1, 26, 6, 30, 6, True: CuadroExcel xlHoja1, 30, 6, 34, 6, True
            CuadroExcel xlHoja1, 34, 6, 38, 6, True: CuadroExcel xlHoja1, 38, 6, 42, 6, True
            CuadroExcel xlHoja1, 42, 6, 46, 6, True: CuadroExcel xlHoja1, 46, 6, 50, 6, True
            CuadroExcel xlHoja1, 50, 6, 54, 6, True: CuadroExcel xlHoja1, 54, 6, 58, 6, True
            CuadroExcel xlHoja1, 58, 6, 62, 6, True
            CuadroExcel xlHoja1, 11, 7, 62, 7
            'PRESUP. MES       MOVIM. MES       DIFER. MES       DIFER. AÑO
            'xlHoja1.Cells(6, 13) = "M   E   S   E   S "
            xlHoja1.Cells(6, 12) = "ENERO": xlHoja1.Cells(6, 16) = "FEBRERO": xlHoja1.Cells(6, 20) = "MARZO": xlHoja1.Cells(6, 24) = "ABRIL": xlHoja1.Cells(6, 28) = "MAYO": xlHoja1.Cells(6, 32) = "JUNIO": xlHoja1.Cells(6, 36) = "JULIO": xlHoja1.Cells(6, 40) = "AGOSTO": xlHoja1.Cells(6, 44) = "SETIEMBRE": xlHoja1.Cells(6, 48) = "OCTUBRE": xlHoja1.Cells(6, 52) = "NOVIEMBRE": xlHoja1.Cells(6, 56) = "DICIEMBRE": xlHoja1.Cells(6, 60) = "TOTALES"
            xlHoja1.Cells(7, 11) = "PRESUP. MES": xlHoja1.Cells(7, 12) = "MOVIM. MES": xlHoja1.Cells(7, 13) = "DIFER. MES": xlHoja1.Cells(7, 14) = "DIFER. AÑO": xlHoja1.Cells(7, 15) = "PRESUP. MES": xlHoja1.Cells(7, 16) = "MOVIM. MES": xlHoja1.Cells(7, 17) = "DIFER. MES": xlHoja1.Cells(7, 18) = "DIFER. AÑO"
            xlHoja1.Cells(7, 19) = "PRESUP. MES": xlHoja1.Cells(7, 20) = "MOVIM. MES": xlHoja1.Cells(7, 21) = "DIFER. MES": xlHoja1.Cells(7, 22) = "DIFER. AÑO": xlHoja1.Cells(7, 23) = "PRESUP. MES": xlHoja1.Cells(7, 24) = "MOVIM. MES": xlHoja1.Cells(7, 25) = "DIFER. MES": xlHoja1.Cells(7, 26) = "DIFER. AÑO"
            xlHoja1.Cells(7, 27) = "PRESUP. MES": xlHoja1.Cells(7, 28) = "MOVIM. MES": xlHoja1.Cells(7, 29) = "DIFER. MES": xlHoja1.Cells(7, 30) = "DIFER. AÑO": xlHoja1.Cells(7, 31) = "PRESUP. MES": xlHoja1.Cells(7, 32) = "MOVIM. MES": xlHoja1.Cells(7, 33) = "DIFER. MES": xlHoja1.Cells(7, 34) = "DIFER. AÑO"
            xlHoja1.Cells(7, 35) = "PRESUP. MES": xlHoja1.Cells(7, 36) = "MOVIM. MES": xlHoja1.Cells(7, 37) = "DIFER. MES": xlHoja1.Cells(7, 38) = "DIFER. AÑO": xlHoja1.Cells(7, 39) = "PRESUP. MES": xlHoja1.Cells(7, 40) = "MOVIM. MES": xlHoja1.Cells(7, 41) = "DIFER. MES": xlHoja1.Cells(7, 42) = "DIFER. AÑO"
            xlHoja1.Cells(7, 43) = "PRESUP. MES": xlHoja1.Cells(7, 44) = "MOVIM. MES": xlHoja1.Cells(7, 45) = "DIFER. MES": xlHoja1.Cells(7, 46) = "DIFER. AÑO": xlHoja1.Cells(7, 47) = "PRESUP. MES": xlHoja1.Cells(7, 48) = "MOVIM. MES": xlHoja1.Cells(7, 49) = "DIFER. MES": xlHoja1.Cells(7, 50) = "DIFER. AÑO"
            xlHoja1.Cells(7, 51) = "PRESUP. MES": xlHoja1.Cells(7, 52) = "MOVIM. MES": xlHoja1.Cells(7, 53) = "DIFER. MES": xlHoja1.Cells(7, 54) = "DIFER. AÑO": xlHoja1.Cells(7, 55) = "PRESUP. MES": xlHoja1.Cells(7, 56) = "MOVIM. MES": xlHoja1.Cells(7, 57) = "DIFER. MES": xlHoja1.Cells(7, 58) = "DIFER. AÑO"
            xlHoja1.Cells(7, 59) = "PRES.TOTAL": xlHoja1.Cells(7, 60) = "PRES.MOVIM.": xlHoja1.Cells(7, 61) = "PRES.DIFER.": xlHoja1.Cells(7, 62) = "AVANCE %"
        ElseIf pnTipo = 2 Then
            'CuadroExcel xlHoja1, 9, 6, 24, 6, True
            CuadroExcel xlHoja1, 11, 6, 14, 6, True: CuadroExcel xlHoja1, 14, 6, 18, 6, True
            CuadroExcel xlHoja1, 18, 6, 22, 6, True: CuadroExcel xlHoja1, 22, 6, 26, 6, True
            CuadroExcel xlHoja1, 26, 6, 30, 6, True
            CuadroExcel xlHoja1, 11, 7, 30, 7
            'xlHoja1.Cells(6, 10) = "T R I M E S T R E S"
            xlHoja1.Cells(6, 12) = "PRIMERO": xlHoja1.Cells(6, 16) = "SEGUNDO": xlHoja1.Cells(6, 20) = "TERCERO": xlHoja1.Cells(6, 24) = "CUARTO": xlHoja1.Cells(6, 28) = "TOTALES"
            xlHoja1.Cells(7, 11) = "PRESUP. MES": xlHoja1.Cells(7, 12) = "MOVIM. MES": xlHoja1.Cells(7, 13) = "DIFER. MES": xlHoja1.Cells(7, 14) = "DIFER. AÑO": xlHoja1.Cells(7, 15) = "PRESUP. MES": xlHoja1.Cells(7, 16) = "MOVIM. MES": xlHoja1.Cells(7, 17) = "DIFER. MES": xlHoja1.Cells(7, 18) = "DIFER. AÑO"
            xlHoja1.Cells(7, 19) = "PRESUP. MES": xlHoja1.Cells(7, 20) = "MOVIM. MES": xlHoja1.Cells(7, 21) = "DIFER. MES": xlHoja1.Cells(7, 22) = "DIFER. AÑO": xlHoja1.Cells(7, 23) = "PRESUP. MES": xlHoja1.Cells(7, 24) = "MOVIM. MES": xlHoja1.Cells(7, 25) = "DIFER. MES": xlHoja1.Cells(7, 26) = "DIFER. AÑO"
            xlHoja1.Cells(7, 27) = "PRES.TOTAL": xlHoja1.Cells(7, 28) = "PRES.MOVIM.": xlHoja1.Cells(7, 29) = "PRES.DIFER.": xlHoja1.Cells(7, 30) = "AVANCE %"
        Else
            'CuadroExcel xlHoja1, 9, 6, 16, 6, True
            CuadroExcel xlHoja1, 11, 6, 14, 6, True: CuadroExcel xlHoja1, 14, 6, 18, 6, True
            CuadroExcel xlHoja1, 18, 6, 22, 6, True
            CuadroExcel xlHoja1, 11, 7, 22, 7
            'xlHoja1.Cells(6, 9) = "S E M E S T R E S"
            xlHoja1.Cells(6, 12) = "PRIMERO": xlHoja1.Cells(6, 16) = "SEGUNDO": xlHoja1.Cells(6, 20) = "TOTALES"
            xlHoja1.Cells(7, 11) = "PRESUP. MES": xlHoja1.Cells(7, 12) = "MOVIM. MES": xlHoja1.Cells(7, 13) = "DIFER. MES": xlHoja1.Cells(7, 14) = "DIFER. AÑO": xlHoja1.Cells(7, 15) = "PRESUP. MES": xlHoja1.Cells(7, 16) = "MOVIM. MES": xlHoja1.Cells(7, 17) = "DIFER. MES": xlHoja1.Cells(7, 18) = "DIFER. AÑO"
            xlHoja1.Cells(7, 19) = "PRES.TOTAL": xlHoja1.Cells(7, 20) = "PRES.MOVIM.": xlHoja1.Cells(7, 21) = "PRES.DIFER.": xlHoja1.Cells(7, 22) = "AVANCE %"
        End If
    End If
    nFil = 7
    nItem = 0
    
    Dim oPP As New DPresupuesto
    'Set tmpReg = oPP.GetPresupuestoEjec(Right(Me.cboPresu.Text, 4), Me.cboFecha.Text, True, chkProy.value, chkBala.value, pnTipo, CInt(Right(Trim(cboTpo.Text), 1)), Right(Me.cboMonedaP.Text, 1), IIf(Me.chkCierreAño.value = 1, True, False))
    Set tmpReg = oPP.GetPresupuestoEjec(Right(Me.cboPresu.Text, 4), Me.cboFecha.Text, True, chkProy.value, chkBala.value, pnTipo, CInt(Right(Trim(cboTpo.Text), 1)), Right(Me.cboMonedaP.Text, 1), IIf(Me.chkCierreAño.value = 1, True, False), Me.chkMovContables.value)
    Set oPP = Nothing
    
    'Carga el Select a ejecutar
    'tmpSql = CargaSelect(pnTipo)
    
    'Set tmpReg = oCon.CargaRecordSet(tmpSql)
    
    If Not (tmpReg.BOF Or tmpReg.EOF) Then
        xlAplicacion.Range("A1:BJ" & Trim(tmpReg.RecordCount + 10)).Font.Size = 7
        With tmpReg
            Do While Not .EOF
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
                    xlHoja1.Cells(nFil, 10) = Format(Round((((!Total / !nMonIni) - 1) * 100), 2), "#,##0.00")
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
                
                If Right(Trim(cboTpo.Text), 1) = "2" Then
                    'SOLO BALANCES
        
                    If .Fields(0) = "00020101010403" Then
                        MsgBox "ddd"
                    End If
        
        
                    nMonBal = !nMonIni
                    nMonMov = !nMonIni
                    nMonSdo = IIf(IsNull(!nSaldoIni), 0, !nSaldoIni)
                    If pnTipo = 1 Then
                        nMonBal = nMonBal + !Ene: nMonMov = nMonMov + !MovEne
                        nMonSdo = nMonSdo + !MovEne
                        xlHoja1.Cells(nFil, 11) = IIf(!Ene = 0, "", Format(nMonBal, "#,##0.00"))
                        xlHoja1.Cells(nFil, 12) = IIf(!MovEne = 0, "", Format(nMonMov, "#,##0.00"))
                        xlHoja1.Cells(nFil, 13) = IIf(nMonSdo = 0, "", Format(nMonSdo, "#,##0.00"))
                        
                        nMonBal = nMonBal + !Feb: nMonMov = nMonMov + !MovFEb
                        nMonSdo = nMonSdo + !MovFEb
                        xlHoja1.Cells(nFil, 14) = IIf(!Feb = 0, "", Format(nMonBal, "#,##0.00"))
                        xlHoja1.Cells(nFil, 15) = IIf(!MovFEb = 0, "", Format(nMonMov, "#,##0.00"))
                        xlHoja1.Cells(nFil, 16) = IIf(nMonSdo = 0, "", Format(nMonSdo, "#,##0.00"))
                        
                        nMonBal = nMonBal + !Mar: nMonMov = nMonMov + !MovMar
                        nMonSdo = nMonSdo + !MovMar
                        xlHoja1.Cells(nFil, 17) = IIf(!Mar = 0, "", Format(nMonBal, "#,##0.00"))
                        xlHoja1.Cells(nFil, 18) = IIf(!MovMar = 0, "", Format(nMonMov, "#,##0.00"))
                        xlHoja1.Cells(nFil, 19) = IIf(nMonSdo = 0, "", Format(nMonSdo, "#,##0.00"))
                        
                        nMonBal = nMonBal + !Abr: nMonMov = nMonMov + !MovAbr
                        nMonSdo = nMonSdo + !MovAbr
                        xlHoja1.Cells(nFil, 20) = IIf(!Abr = 0, "", Format(nMonBal, "#,##0.00"))
                        xlHoja1.Cells(nFil, 21) = IIf(!MovAbr = 0, "", Format(nMonMov, "#,##0.00"))
                        xlHoja1.Cells(nFil, 22) = IIf(nMonSdo = 0, "", Format(nMonSdo, "#,##0.00"))
                        
                        nMonBal = nMonBal + !May: nMonMov = nMonMov + !MovMay
                        nMonSdo = nMonSdo + !MovMay
                        xlHoja1.Cells(nFil, 23) = IIf(!May = 0, "", Format(nMonBal, "#,##0.00"))
                        xlHoja1.Cells(nFil, 24) = IIf(!MovMay = 0, "", Format(nMonMov, "#,##0.00"))
                        xlHoja1.Cells(nFil, 25) = IIf(nMonSdo = 0, "", Format(nMonSdo, "#,##0.00"))
                        
                        nMonBal = nMonBal + !Jun: nMonMov = nMonMov + !MovJun
                        nMonSdo = nMonSdo + !MovJun
                        xlHoja1.Cells(nFil, 26) = IIf(!Jun = 0, "", Format(nMonBal, "#,##0.00"))
                        xlHoja1.Cells(nFil, 27) = IIf(!MovJun = 0, "", Format(nMonMov, "#,##0.00"))
                        xlHoja1.Cells(nFil, 28) = IIf(nMonSdo = 0, "", Format(nMonSdo, "#,##0.00"))
                        
                        nMonBal = nMonBal + !Jul: nMonMov = nMonMov + !MovJul
                        nMonSdo = nMonSdo + !MovJul
                        xlHoja1.Cells(nFil, 29) = IIf(!Jul = 0, "", Format(nMonBal, "#,##0.00"))
                        xlHoja1.Cells(nFil, 30) = IIf(!MovJul = 0, "", Format(nMonMov, "#,##0.00"))
                        xlHoja1.Cells(nFil, 31) = IIf(nMonSdo = 0, "", Format(nMonSdo, "#,##0.00"))
                        
                        nMonBal = nMonBal + !Ago: nMonMov = nMonMov + !MovAgo
                        nMonSdo = nMonSdo + !MovAgo
                        xlHoja1.Cells(nFil, 32) = IIf(!Ago = 0, "", Format(nMonBal, "#,##0.00"))
                        xlHoja1.Cells(nFil, 33) = IIf(!MovAgo = 0, "", Format(nMonMov, "#,##0.00"))
                        xlHoja1.Cells(nFil, 34) = IIf(nMonSdo = 0, "", Format(nMonSdo, "#,##0.00"))
                        
                        nMonBal = nMonBal + !Seti: nMonMov = nMonMov + !MovSet
                        nMonSdo = nMonSdo + !MovSet
                        xlHoja1.Cells(nFil, 35) = IIf(!Seti = 0, "", Format(nMonBal, "#,##0.00"))
                        xlHoja1.Cells(nFil, 36) = IIf(!MovSet = 0, "", Format(nMonMov, "#,##0.00"))
                        xlHoja1.Cells(nFil, 37) = IIf(nMonSdo = 0, "", Format(nMonSdo, "#,##0.00"))
                        
                        nMonBal = nMonBal + !Oct: nMonMov = nMonMov + !MovOct
                        nMonSdo = nMonSdo + !MovOct
                        xlHoja1.Cells(nFil, 38) = IIf(!Oct = 0, "", Format(nMonBal, "#,##0.00"))
                        xlHoja1.Cells(nFil, 39) = IIf(!MovOct = 0, "", Format(nMonMov, "#,##0.00"))
                        xlHoja1.Cells(nFil, 40) = IIf(nMonSdo = 0, "", Format(nMonSdo, "#,##0.00"))
                        
                        nMonBal = nMonBal + !Nov: nMonMov = nMonMov + !MovNov
                        nMonSdo = nMonSdo + !MovNov
                        xlHoja1.Cells(nFil, 41) = IIf(!Nov = 0, "", Format(nMonBal, "#,##0.00"))
                        xlHoja1.Cells(nFil, 42) = IIf(!MovNov = 0, "", Format(nMonMov, "#,##0.00"))
                        xlHoja1.Cells(nFil, 43) = IIf(nMonSdo = 0, "", Format(nMonSdo, "#,##0.00"))
                        
                        nMonBal = nMonBal + !Dic: nMonMov = nMonMov + !MovDic
                        nMonSdo = nMonSdo + !MovDic
                        xlHoja1.Cells(nFil, 44) = IIf(!Dic = 0, "", Format(nMonBal, "#,##0.00"))
                        xlHoja1.Cells(nFil, 45) = IIf(!MovDic = 0, "", Format(nMonMov, "#,##0.00"))
                        xlHoja1.Cells(nFil, 46) = IIf(nMonSdo = 0, "", Format(nMonSdo, "#,##0.00"))
                    
                        If nMonBal > 0 And nMonMov > 0 Then
                            xlHoja1.Cells(nFil, 47) = Round(((nMonMov / nMonBal) * 100), 2)
                        Else
                            xlHoja1.Cells(nFil, 47) = ""
                        End If
                    ElseIf pnTipo = 2 Then
                        nMonBal = nMonBal + !Primero: nMonMov = nMonMov + !MovPrimero
                        nMonSdo = nMonSdo + !MovPrimero
                        xlHoja1.Cells(nFil, 11) = IIf(!Primero = 0, "", Format(nMonBal, "#,##0.00"))
                        xlHoja1.Cells(nFil, 12) = IIf(!MovPrimero = 0, "", Format(nMonMov, "#,##0.00"))
                        xlHoja1.Cells(nFil, 13) = IIf(nMonSdo = 0, "", Format(nMonSdo, "#,##0.00"))
                        
                        nMonBal = nMonBal + !Segundo: nMonMov = nMonMov + !MovSegundo
                        nMonSdo = nMonSdo + !MovSegundo
                        xlHoja1.Cells(nFil, 14) = IIf(!Segundo = 0, "", Format(nMonBal, "#,##0.00"))
                        xlHoja1.Cells(nFil, 15) = IIf(!MovSegundo = 0, "", Format(nMonMov, "#,##0.00"))
                        xlHoja1.Cells(nFil, 16) = IIf(nMonSdo = 0, "", Format(nMonSdo, "#,##0.00"))
                        
                        nMonBal = nMonBal + !Tercero: nMonMov = nMonMov + !MovTercero
                        nMonSdo = nMonSdo + !MovTercero
                        xlHoja1.Cells(nFil, 17) = IIf(!Tercero = 0, "", Format(nMonBal, "#,##0.00"))
                        xlHoja1.Cells(nFil, 18) = IIf(!MovTercero = 0, "", Format(nMonMov, "#,##0.00"))
                        xlHoja1.Cells(nFil, 19) = IIf(nMonSdo = 0, "", Format(nMonSdo, "#,##0.00"))
                        
                        nMonBal = nMonBal + !Cuarto: nMonMov = nMonMov + !MovCuarto
                        nMonSdo = nMonSdo + !MovCuarto
                        xlHoja1.Cells(nFil, 20) = IIf(!Cuarto = 0, "", Format(nMonBal, "#,##0.00"))
                        xlHoja1.Cells(nFil, 21) = IIf(!MovCuarto = 0, "", Format(nMonMov, "#,##0.00"))
                        xlHoja1.Cells(nFil, 22) = IIf(nMonSdo = 0, "", Format(nMonSdo, "#,##0.00"))
                        
                        If nMonBal > 0 And nMonMov > 0 Then
                            xlHoja1.Cells(nFil, 23) = Round(((nMonMov / nMonBal) * 100), 2)
                        Else
                            xlHoja1.Cells(nFil, 23) = ""
                        End If
                    Else
                        nMonBal = nMonBal + !Primero: nMonMov = nMonMov + !MovPrimero
                        nMonSdo = nMonSdo + !MovPrimero
                        xlHoja1.Cells(nFil, 11) = IIf(!Primero = 0, "", Format(nMonBal, "#,##0.00"))
                        xlHoja1.Cells(nFil, 12) = IIf(!MovPrimero = 0, "", Format(nMonMov, "#,##0.00"))
                        xlHoja1.Cells(nFil, 13) = IIf(nMonSdo = 0, "", Format(nMonSdo, "#,##0.00"))
                        
                        nMonBal = nMonBal + !Segundo: nMonMov = nMonMov + !MovSegundo
                        nMonSdo = nMonSdo + !MovSegundo
                        xlHoja1.Cells(nFil, 14) = IIf(!Segundo = 0, "", Format(nMonBal, "#,##0.00"))
                        xlHoja1.Cells(nFil, 15) = IIf(!MovSegundo = 0, "", Format(nMonMov, "#,##0.00"))
                        xlHoja1.Cells(nFil, 16) = IIf(nMonSdo = 0, "", Format(nMonSdo, "#,##0.00"))
                        
                        If nMonBal > 0 And nMonMov > 0 Then
                            xlHoja1.Cells(nFil, 17) = Round(((nMonMov / nMonBal) * 100), 2)
                        Else
                            xlHoja1.Cells(nFil, 17) = ""
                        End If
                    End If
                Else
                    If pnTipo = 1 Then
                        xlHoja1.Cells(nFil, 11) = IIf(!Ene = 0, "", Format(!Ene, "#,##0.00"))
                        xlHoja1.Cells(nFil, 12) = IIf(!MovEne = 0, "", Format(!MovEne, "#,##0.00"))
                        xlHoja1.Cells(nFil, 13) = IIf(!DMEne = 0, "", Format(!DMEne, "#,##0.00"))
                        xlHoja1.Cells(nFil, 14) = IIf(!DMEne = 0, "", Format(!DMEne, "#,##0.00"))
                        xlHoja1.Cells(nFil, 15) = IIf(!Feb = 0, "", Format(!Feb, "#,##0.00"))
                        xlHoja1.Cells(nFil, 16) = IIf(!MovFEb = 0, "", Format(!MovFEb, "#,##0.00"))
                        xlHoja1.Cells(nFil, 17) = IIf(!DMFeb = 0, "", Format(!DMFeb, "#,##0.00"))
                        xlHoja1.Cells(nFil, 18) = IIf(!DMEne + !DMFeb = 0, "", Format(!DMEne + !DMFeb, "#,##0.00"))
                        xlHoja1.Cells(nFil, 19) = IIf(!Mar = 0, "", Format(!Mar, "#,##0.00"))
                        xlHoja1.Cells(nFil, 20) = IIf(!MovMar = 0, "", Format(!MovMar, "#,##0.00"))
                        xlHoja1.Cells(nFil, 21) = IIf(!DMMar = 0, "", Format(!DMMar, "#,##0.00"))
                        xlHoja1.Cells(nFil, 22) = IIf(!DMEne + !DMFeb + !DMMar = 0, "", Format(!DMEne + !DMFeb + !DMMar, "#,##0.00"))
                        xlHoja1.Cells(nFil, 23) = IIf(!Abr = 0, "", Format(!Abr, "#,##0.00"))
                        xlHoja1.Cells(nFil, 24) = IIf(!MovAbr = 0, "", Format(!MovAbr, "#,##0.00"))
                        xlHoja1.Cells(nFil, 25) = IIf(!DMAbr = 0, "", Format(!DMAbr, "#,##0.00"))
                        xlHoja1.Cells(nFil, 26) = IIf(!DMEne + !DMFeb + !DMMar + !DMAbr = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr, "#,##0.00"))
                        xlHoja1.Cells(nFil, 27) = IIf(!May = 0, "", Format(!May, "#,##0.00"))
                        xlHoja1.Cells(nFil, 28) = IIf(!MovMay = 0, "", Format(!MovMay, "#,##0.00"))
                        xlHoja1.Cells(nFil, 29) = IIf(!DMMay = 0, "", Format(!DMMay, "#,##0.00"))
                        xlHoja1.Cells(nFil, 30) = IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay, "#,##0.00"))
                        xlHoja1.Cells(nFil, 31) = IIf(!Jun = 0, "", Format(!Jun, "#,##0.00"))
                        xlHoja1.Cells(nFil, 32) = IIf(!MovJun = 0, "", Format(!MovJun, "#,##0.00"))
                        xlHoja1.Cells(nFil, 33) = IIf(!DMJun = 0, "", Format(!DMJun, "#,##0.00"))
                        xlHoja1.Cells(nFil, 34) = IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun, "#,##0.00"))
                        xlHoja1.Cells(nFil, 35) = IIf(!Jul = 0, "", Format(!Jul, "#,##0.00"))
                        xlHoja1.Cells(nFil, 36) = IIf(!MovJul = 0, "", Format(!MovJul, "#,##0.00"))
                        xlHoja1.Cells(nFil, 37) = IIf(!DMJul = 0, "", Format(!DMJul, "#,##0.00"))
                        xlHoja1.Cells(nFil, 38) = IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul, "#,##0.00"))
                        xlHoja1.Cells(nFil, 39) = IIf(!Ago = 0, "", Format(!Ago, "#,##0.00"))
                        xlHoja1.Cells(nFil, 40) = IIf(!MovAgo = 0, "", Format(!MovAgo, "#,##0.00"))
                        xlHoja1.Cells(nFil, 41) = IIf(!DMAgo = 0, "", Format(!DMAgo, "#,##0.00"))
                        xlHoja1.Cells(nFil, 42) = IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo, "#,##0.00"))
                        xlHoja1.Cells(nFil, 43) = IIf(!Seti = 0, "", Format(!Seti, "#,##0.00"))
                        xlHoja1.Cells(nFil, 44) = IIf(!MovSet = 0, "", Format(!MovSet, "#,##0.00"))
                        xlHoja1.Cells(nFil, 45) = IIf(!DMSet = 0, "", Format(!DMSet, "#,##0.00"))
                        xlHoja1.Cells(nFil, 46) = IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet, "#,##0.00"))
                        xlHoja1.Cells(nFil, 47) = IIf(!Oct = 0, "", Format(!Oct, "#,##0.00"))
                        xlHoja1.Cells(nFil, 48) = IIf(!MovOct = 0, "", Format(!MovOct, "#,##0.00"))
                        xlHoja1.Cells(nFil, 49) = IIf(!DMOct = 0, "", Format(!DMOct, "#,##0.00"))
                        xlHoja1.Cells(nFil, 50) = IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet + !DMOct = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet + !DMOct, "#,##0.00"))
                        xlHoja1.Cells(nFil, 51) = IIf(!Nov = 0, "", Format(!Nov, "#,##0.00"))
                        xlHoja1.Cells(nFil, 52) = IIf(!MovNov = 0, "", Format(!MovNov, "#,##0.00"))
                        xlHoja1.Cells(nFil, 53) = IIf(!DMNov = 0, "", Format(!DMNov, "#,##0.00"))
                        xlHoja1.Cells(nFil, 54) = IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet + !DMOct + !DMNov = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet + !DMOct + !DMNov, "#,##0.00"))
                        xlHoja1.Cells(nFil, 55) = IIf(!Dic = 0, "", Format(!Dic, "#,##0.00"))
                        xlHoja1.Cells(nFil, 56) = IIf(!MovDic = 0, "", Format(!MovDic, "#,##0.00"))
                        xlHoja1.Cells(nFil, 57) = IIf(!DMDic = 0, "", Format(!DMDic, "#,##0.00"))
                        xlHoja1.Cells(nFil, 58) = IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet + !DMOct + !DMNov + !DMDic = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet + !DMOct + !DMNov + !DMDic, "#,##0.00"))
                
                        'Totales
                        TotMov = !MovEne + !MovFEb + !MovMar + !MovAbr + !MovMay + !MovJun + !MovJul + !MovAgo + !MovSet + !MovOct + !MovNov + !MovDic
                        xlHoja1.Cells(nFil, 59) = IIf(!Total = 0, "", Format(!Total, "#,##0.00"))
                        xlHoja1.Cells(nFil, 60) = IIf(TotMov = 0, "", Format(TotMov, "#,##0.00"))
                        xlHoja1.Cells(nFil, 61) = IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet + !DMOct + !DMNov + !DMDic = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet + !DMOct + !DMNov + !DMDic, "#,##0.00"))
                        If !Total > 0 And TotMov > 0 Then
                            xlHoja1.Cells(nFil, 62) = Round(((TotMov / !Total) * 100), 2)
                        Else
                            xlHoja1.Cells(nFil, 62) = ""
                        End If
                    
                    ElseIf pnTipo = 2 Then
                        xlHoja1.Cells(nFil, 11) = IIf(!Primero = 0, "", Format(!Primero, "#,##0.00"))
                        xlHoja1.Cells(nFil, 12) = IIf(!MovPrimero = 0, "", Format(!MovPrimero, "#,##0.00"))
                        xlHoja1.Cells(nFil, 13) = IIf(!DMPrimero = 0, "", Format(!DMPrimero, "#,##0.00"))
                        xlHoja1.Cells(nFil, 14) = IIf(!DMPrimero = 0, "", Format(!DMPrimero, "#,##0.00"))
                        
                        xlHoja1.Cells(nFil, 15) = IIf(!Segundo = 0, "", Format(!Segundo, "#,##0.00"))
                        xlHoja1.Cells(nFil, 16) = IIf(!MovSegundo = 0, "", Format(!MovSegundo, "#,##0.00"))
                        xlHoja1.Cells(nFil, 17) = IIf(!DMSegundo = 0, "", Format(!DMSegundo, "#,##0.00"))
                        xlHoja1.Cells(nFil, 18) = IIf(!DMPrimero + !DMSegundo = 0, "", Format(!DMPrimero + !DMSegundo, "#,##0.00"))
                        
                        xlHoja1.Cells(nFil, 19) = IIf(!Tercero = 0, "", Format(!Tercero, "#,##0.00"))
                        xlHoja1.Cells(nFil, 20) = IIf(!MovTercero = 0, "", Format(!MovTercero, "#,##0.00"))
                        xlHoja1.Cells(nFil, 21) = IIf(!DMTercero = 0, "", Format(!DMTercero, "#,##0.00"))
                        xlHoja1.Cells(nFil, 22) = IIf(!DMPrimero + !DMSegundo + !DMTercero = 0, "", Format(!DMPrimero + !DMSegundo + !DMTercero, "#,##0.00"))
                        
                        xlHoja1.Cells(nFil, 23) = IIf(!Cuarto = 0, "", Format(!Cuarto, "#,##0.00"))
                        xlHoja1.Cells(nFil, 24) = IIf(!MovCuarto = 0, "", Format(!MovCuarto, "#,##0.00"))
                        xlHoja1.Cells(nFil, 25) = IIf(!DMCuarto = 0, "", Format(!DMCuarto, "#,##0.00"))
                        xlHoja1.Cells(nFil, 26) = IIf(!DMPrimero + !DMSegundo + !DMTercero + !DMCuarto = 0, "", Format(!DMPrimero + !DMSegundo + !DMTercero + !DMCuarto, "#,##0.00"))
                    
                        'Totales
                        TotMov = !MovPrimero + !MovSegundo + !MovTercero + !MovCuarto
                        xlHoja1.Cells(nFil, 27) = IIf(!Total = 0, "", Format(!Total, "#,##0.00"))
                        xlHoja1.Cells(nFil, 28) = IIf(TotMov = 0, "", Format(TotMov, "#,##0.00"))
                        xlHoja1.Cells(nFil, 29) = IIf(!DMPrimero + !DMSegundo + !DMTercero + !DMCuarto = 0, "", Format(!DMPrimero + !DMSegundo + !DMTercero + !DMCuarto, "#,##0.00"))
                        If !Total > 0 And TotMov > 0 Then
                            xlHoja1.Cells(nFil, 30) = Round(((TotMov / !Total) * 100), 2)
                        Else
                            xlHoja1.Cells(nFil, 30) = ""
                        End If
                    Else
                        xlHoja1.Cells(nFil, 11) = IIf(!Primero = 0, "", Format(!Primero, "#,##0.00"))
                        xlHoja1.Cells(nFil, 12) = IIf(!MovPrimero = 0, "", Format(!MovPrimero, "#,##0.00"))
                        xlHoja1.Cells(nFil, 13) = IIf(!DMPrimero = 0, "", Format(!DMPrimero, "#,##0.00"))
                        xlHoja1.Cells(nFil, 14) = IIf(!DMPrimero = 0, "", Format(!DMPrimero, "#,##0.00"))
                        
                        xlHoja1.Cells(nFil, 15) = IIf(!Segundo = 0, "", Format(!Segundo, "#,##0.00"))
                        xlHoja1.Cells(nFil, 16) = IIf(!MovSegundo = 0, "", Format(!MovSegundo, "#,##0.00"))
                        xlHoja1.Cells(nFil, 17) = IIf(!DMSegundo = 0, "", Format(!DMSegundo, "#,##0.00"))
                        xlHoja1.Cells(nFil, 18) = IIf(!DMPrimero + !DMSegundo = 0, "", Format(!DMPrimero + !DMSegundo, "#,##0.00"))
                        
                        'Totales
                        TotMov = !MovPrimero + !MovSegundo
                        xlHoja1.Cells(nFil, 19) = IIf(!Total = 0, "", Format(!Total, "#,##0.00"))
                        xlHoja1.Cells(nFil, 20) = IIf(TotMov = 0, "", Format(TotMov, "#,##0.00"))
                        xlHoja1.Cells(nFil, 21) = IIf(!DMPrimero + !DMSegundo = 0, "", Format(!DMPrimero + !DMSegundo, "#,##0.00"))
                        If !Total > 0 And TotMov > 0 Then
                            xlHoja1.Cells(nFil, 22) = Round(((TotMov / !Total) * 100), 2)
                        Else
                            xlHoja1.Cells(nFil, 22) = ""
                        End If
                    End If
                End If
                .MoveNext
            Loop
        End With
    End If
    tmpReg.Close
    Set tmpReg = Nothing
    If Right(Trim(cboTpo.Text), 1) = "2" Then
        'SOLO BALANCES
        If pnTipo = 1 Then
            CuadroExcel xlHoja1, 1, 8, 47, nFil
        ElseIf pnTipo = 2 Then
            CuadroExcel xlHoja1, 1, 8, 23, nFil
        Else
            CuadroExcel xlHoja1, 1, 8, 17, nFil
        End If
    Else
        If pnTipo = 1 Then
            CuadroExcel xlHoja1, 1, 8, 62, nFil
        ElseIf pnTipo = 2 Then
            CuadroExcel xlHoja1, 1, 8, 30, nFil
        Else
            CuadroExcel xlHoja1, 1, 8, 22, nFil
        End If
    End If
    ExcelEnd App.path & "\SPOOLER\" & psArchivo, xlAplicacion, xlLibro, xlHoja1
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

Private Function CargaSelect(ByVal pnTipo As Integer) As String
    Dim tmpSql  As String
    Dim lsCadena As String
    If pnTipo = 1 Then 'Mensual
        tmpSql = "SELECT rr.cCodRub, rr.cDesRub, sum(res.nMonIni) nMonIni, sum(res.nMonto) nMonto, sum(res.nMonCre1) nMonCre1, sum(res.nMonCre2) nMonCre2, " & _
            "       SUM(res.nMonto + res.nMonCre1 + res.nMonCre2) Total, " & _
            "       SUM(Ene) Ene, SUM(MovEne) MovEne, SUM(DMEne) DMEne, " & _
            "       SUM(Feb) Feb, SUM(MovFeb) MovFeb, SUM(DMFeb) DMFeb, " & _
            "       SUM(Mar) Mar, SUM(MovMar) MovMar, SUM(DMMar) DMMar, " & _
            "       SUM(Abr) Abr, SUM(MovAbr) MovAbr, SUM(DMAbr) DMAbr, " & _
            "       SUM(May) May, SUM(MovMay) MovMay, SUM(DMMay) DMMay, " & _
            "       SUM(Jun) Jun, SUM(MovJun) MovJun, SUM(DMJun) DMJun, " & _
            "       SUM(Jul) Jul, SUM(MovJul) MovJul, SUM(DMJul) DMJul, " & _
            "       SUM(Ago) Ago, SUM(MovAgo) MovAgo, SUM(DMAgo) DMAgo, " & _
            "       SUM(Seti) Seti, SUM(MovSet) MovSet, SUM(DMSet) DMSet, " & _
            "       SUM(Oct) Oct, SUM(MovOct) MovOct, SUM(DMOct) DMOct, " & _
            "       SUM(Nov) Nov, SUM(MovNov) MovNov, SUM(DMNov) DMNov, " & _
            "       SUM(Dic) Dic, SUM(MovDic) MovDic, SUM(DMDic) DMDic " & _
            " FROM (SELECT DISTINCT cCodRub, cDesRub " & _
            "       FROM pRubro " & _
            "       WHERE cAno = '" & Left(Trim(cboFecha.Text), 4) & "' AND cPresu = '" & Right(cboPresu.Text, 4) & "' " & _
            "       ) RR " & _
            "           LEFT JOIN " & _
            "              (    "
        If chkProy.value = 1 Then
            'Movimientos CON Proyeccion
            tmpSql = ""
            tmpSql = tmpSql + " SELECT r.cCodRub, r.cDesRub, r.nMonIni, r.nMonto, r.nMonCre1, r.nMonCre2, " & _
                "                   (r.nMonto + r.nMonCre1 + r.nMonCre2) Total ," & _
                "                   sum(CASE WHEN p.cPeriodo = 'P301' THEN p.nMonIni ELSE 0 END) Ene," & _
                "                   (IsNull(Cta.MovEne,0) + IsNull(Pro.ProEne,0)) MovEne, " & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P301' THEN p.nMonIni ELSE 0 END)" & _
                "                   - (IsNull(Cta.MovEne,0) + IsNull(Pro.ProEne,0))) DMEne, " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P302' THEN p.nMonIni ELSE 0 END) Feb," & _
                "                   (IsNull(Cta.MovFeb,0) + IsNull(Pro.ProFeb,0)) MovFeb, " & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P302' THEN p.nMonIni ELSE 0 END)" & _
                "                   - (IsNull(Cta.MovFeb,0) + + IsNull(Pro.ProFeb,0))) DMFeb, " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P303' THEN p.nMonIni ELSE 0 END) Mar," & _
                "                   (IsNull(Cta.MovMar,0) + IsNull(Pro.ProMar,0)) MovMar, " & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P303' THEN p.nMonIni ELSE 0 END)" & _
                "                   - (IsNull(Cta.MovMar,0) + IsNull(Pro.ProMar,0))) DMMar, " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P304' THEN p.nMonIni ELSE 0 END) Abr," & _
                "                   (IsNull(Cta.MovAbr,0) + IsNull(Pro.ProAbr,0)) MovAbr," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P304' THEN p.nMonIni ELSE 0 END)" & _
                "                   - (IsNull(Cta.MovAbr,0) + IsNull(Pro.ProAbr,0))) DMAbr, "
            tmpSql = tmpSql & _
                "                   sum(CASE WHEN p.cPeriodo = 'P305' THEN p.nMonIni ELSE 0 END) May," & _
                "                   (IsNull(Cta.MovMay,0) + IsNull(Pro.ProMay,0)) MovMay," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P305' THEN p.nMonIni ELSE 0 END)" & _
                "                   - (IsNull(Cta.MovMay,0) + IsNull(Pro.ProMay,0))) DMMay, " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P306' THEN p.nMonIni ELSE 0 END) Jun," & _
                "                   (IsNull(Cta.MovJun,0) + IsNull(Pro.ProJun,0)) MovJun," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P306' THEN p.nMonIni ELSE 0 END)" & _
                "                   - (IsNull(Cta.MovJun,0) + IsNull(Pro.ProJun,0))) DMJun, " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P307' THEN p.nMonIni ELSE 0 END) Jul," & _
                "                   (IsNull(Cta.MovJul,0) + IsNull(Pro.ProJul,0)) MovJul," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P307' THEN p.nMonIni ELSE 0 END)" & _
                "                   - (IsNull(Cta.MovJul,0) + IsNull(Pro.ProJul,0))) DMJul, " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P308' THEN p.nMonIni ELSE 0 END) Ago," & _
                "                   (IsNull(Cta.MovAgo,0) + IsNull(Pro.ProAgo,0)) MovAgo," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P308' THEN p.nMonIni ELSE 0 END)" & _
                "                   - (IsNull(Cta.MovAgo,0) + IsNull(Pro.ProAgo,0))) DMAgo, "
            tmpSql = tmpSql & _
                "                   sum(CASE WHEN p.cPeriodo = 'P309' THEN p.nMonIni ELSE 0 END) Seti," & _
                "                   (IsNull(Cta.MovSet,0) + IsNull(Pro.ProSet,0)) MovSet," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P309' THEN p.nMonIni ELSE 0 END)" & _
                "                   - (IsNull(Cta.MovSet,0) + IsNull(Pro.ProSet,0))) DMSet, " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P310' THEN p.nMonIni ELSE 0 END) Oct," & _
                "                   (IsNull(Cta.MovOct,0) + IsNull(Pro.ProOct,0)) MovOct," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P310' THEN p.nMonIni ELSE 0 END)" & _
                "                   - (IsNull(Cta.MovOct,0) + IsNull(Pro.ProOct,0))) DMOct, " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P311' THEN p.nMonIni ELSE 0 END) Nov," & _
                "                   (IsNull(Cta.MovNov,0) + IsNull(Pro.ProNov,0)) MovNov," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P311' THEN p.nMonIni ELSE 0 END)" & _
                "                   - (IsNull(Cta.MovNov,0) + IsNull(Pro.ProNov,0))) DMNov, " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P312' THEN p.nMonIni ELSE 0 END) Dic," & _
                "                   (IsNull(Cta.MovDic,0) + IsNull(Pro.ProDic,0)) MovDic," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P312' THEN p.nMonIni ELSE 0 END)" & _
                "                   - (IsNull(Cta.MovDic,0) + IsNull(Pro.ProDic,0))) DMDic "
            tmpSql = tmpSql & " FROM PRubro R LEFT JOIN pPresupu P ON r.cAno = p.cAno AND p.cPresu = r.cPresu AND p.cCodRub = r.cCodRub" & _
                "                   LEFT JOIN " & _
                "                       (SELECT rc.cAno, rc.cPresu, rc.cCodRub, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____01%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovEne, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____02%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovFeb, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____03%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovMar, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____04%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovAbr, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____05%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovMay, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____06%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovJun, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____07%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovJul, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____08%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovAgo, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____09%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovSet, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____10%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovOct, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____11%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovNov, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____12%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovDic " & _
                "                       FROM PRubCta RC JOIN MovCta MC ON SubString(mc.cCtaContCod,1,2) + '_' + SubString(mc.cCtaContCod,4,len(mc.cCtaContCod)) like SubString(rc.cCtaCnt,1,2) + '_' + SubString(rc.cCtaCnt,4,len(rc.cCtaCnt)) + '%' " & _
                "                           JOIN MOV M ON mc.cMovNro = m.cMovNro "
            tmpSql = tmpSql & _
                "                           JOIN (SELECT cValor, cAbrev FROM " & gcCentralCom & "TablaCod " & _
                "                                   WHERE cCodTab LIKE 'C0__') Tip " & _
                "                           ON substring(mc.cCtaContCod,1,len(tip.cValor)) = Tip.cValor "
            tmpSql = tmpSql & _
                "                       WHERE mc.cMovNro LIKE '" & Left(Trim(cboFecha.Text), 4) & "%' " & _
                "                           AND SubString(mc.cCtaContCod,3,1) IN ('1','2') " & _
                "                           AND m.cMovFlag <> 'X' AND (m.cMovEstado='9' OR " & _
                "                                   (m.cMovEstado='0' AND Not Exists(Select Distinct m2.cmovestado " & _
                "                                                           From mov m2 join movref mr " & _
                "                                                               On m2.cmovnro = mr.cmovnroref " & _
                "                                                               And mr.cmovnro = mc.cmovnro " & _
                "                                                           Where m2.cMovFlag <> 'X' And m2.cMovEstado = '9'))) " & _
                "                       GROUP BY rc.cAno, rc.cPresu, rc.cCodRub, tip.cAbrev ) Cta " & _
                "                   ON r.cAno = cta.cAno AND r.cPresu = cta.cPresu AND r.cCodRub = cta.cCodRub "
            tmpSql = tmpSql & _
                "                   LEFT JOIN " & _
                "                       (SELECT  rc.cAno, rc.cPresu, rc.cCodRub, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 1 THEN (msd.nServImporte) ELSE 0 END),0)) ProEne, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 2 THEN (msd.nServImporte) ELSE 0 END),0)) ProFeb, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 3 THEN (msd.nServImporte) ELSE 0 END),0)) ProMar, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 4 THEN (msd.nServImporte) ELSE 0 END),0)) ProAbr, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 5 THEN (msd.nServImporte) ELSE 0 END),0)) ProMay, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 6 THEN (msd.nServImporte) ELSE 0 END),0)) ProJun, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 7 THEN (msd.nServImporte) ELSE 0 END),0)) ProJul, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 8 THEN (msd.nServImporte) ELSE 0 END),0)) ProAgo, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 9 THEN (msd.nServImporte) ELSE 0 END),0)) ProSet, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 10 THEN (msd.nServImporte) ELSE 0 END),0)) ProOct, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 11 THEN (msd.nServImporte) ELSE 0 END),0)) ProNov, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 12 THEN (msd.nServImporte) ELSE 0 END),0)) ProDic " & _
                "                       FROM MovServicioDet MSD JOIN MovServicio MS ON msd.cMovNro = ms.cMovNro " & _
                "                           JOIN Mov M ON ms.cMovNro = m.cMovNro JOIN MovCta MC ON m.cMovNro = mc.cMovNro " & _
                "                           JOIN PRubCta RC ON mc.cCtaContCod = rc.cCtaCnt  AND rc.cPresu = ms.cPresuCod " & _
                "                       WHERE  Year(dServFecha) = '" & Left(Trim(cboFecha.Text), 4) & "' AND month(dServFecha) >= " & Month(gdFecSis) & " " & _
                "                           AND m.cMovEstado='S' AND m.cMovFlag <> 'X' " & _
                "                       GROUP BY rc.cAno, rc.cPresu, rc.cCodRub ) Pro " & _
                "                   ON r.cAno = pro.cAno AND r.cPresu = pro.cPresu AND r.cCodRub = pro.cCodRub "
            tmpSql = tmpSql & _
                "               WHERE r.cAno = '" & Left(Trim(cboFecha.Text), 4) & "' AND r.cPresu = '" & Right(cboPresu.Text, 4) & "'" & _
                "                   AND r.cFlag Is Null" & _
                "               GROUP BY r.cCodRub, r.cDesRub, r.nMonIni, r.nMonto, r.nMonCre1, r.nMonCre2, " & _
                "                   Cta.MovEne, Cta.MovFeb, Cta.MovMar, Cta.MovAbr, Cta.MovMay, Cta.MovJun, " & _
                "                   Cta.MovJul, Cta.MovAgo, Cta.MovSet, Cta.MovOct, Cta.MovNov, Cta.MovDic, " & _
                "                   Pro.ProEne, Pro.ProFeb, Pro.ProMar, Pro.ProAbr, Pro.ProMay, Pro.ProJun, " & _
                "                   Pro.ProJul, Pro.ProAgo, Pro.ProSet, Pro.ProOct, Pro.ProNov, Pro.ProDic "
                
        Else
            'Movimientos SIN Proyección
            tmpSql = tmpSql + " SELECT r.cCodRub, r.cDesRub, r.nMonIni, r.nMonto, r.nMonCre1, r.nMonCre2, " & _
                "                   (r.nMonto + r.nMonCre1 + r.nMonCre2) Total ," & _
                "                   sum(CASE WHEN p.cPeriodo = 'P301' THEN p.nMonIni ELSE 0 END) Ene," & _
                "                   IsNull(Cta.MovEne,0) MovEne, " & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P301' THEN p.nMonIni ELSE 0 END)" & _
                "                   - IsNull(Cta.MovEne,0)) DMEne, " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P302' THEN p.nMonIni ELSE 0 END) Feb," & _
                "                   IsNull(Cta.MovFeb,0) MovFeb, " & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P302' THEN p.nMonIni ELSE 0 END)" & _
                "                   - IsNull(Cta.MovFeb,0)) DMFeb, " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P303' THEN p.nMonIni ELSE 0 END) Mar," & _
                "                   IsNull(Cta.MovMar,0) MovMar, " & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P303' THEN p.nMonIni ELSE 0 END)" & _
                "                   - IsNull(Cta.MovMar,0)) DMMar, " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P304' THEN p.nMonIni ELSE 0 END) Abr," & _
                "                   IsNull(Cta.MovAbr,0) MovAbr," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P304' THEN p.nMonIni ELSE 0 END)" & _
                "                   - IsNull(Cta.MovAbr,0)) DMAbr, "
            tmpSql = tmpSql & _
                "                   sum(CASE WHEN p.cPeriodo = 'P305' THEN p.nMonIni ELSE 0 END) May," & _
                "                   IsNull(Cta.MovMay,0) MovMay," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P305' THEN p.nMonIni ELSE 0 END)" & _
                "                   - IsNull(Cta.MovMay,0)) DMMay, " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P306' THEN p.nMonIni ELSE 0 END) Jun," & _
                "                   IsNull(Cta.MovJun,0) MovJun," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P306' THEN p.nMonIni ELSE 0 END)" & _
                "                   - IsNull(Cta.MovJun,0)) DMJun, " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P307' THEN p.nMonIni ELSE 0 END) Jul," & _
                "                   IsNull(Cta.MovJul,0) MovJul," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P307' THEN p.nMonIni ELSE 0 END)" & _
                "                   - IsNull(Cta.MovJul,0)) DMJul, " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P308' THEN p.nMonIni ELSE 0 END) Ago," & _
                "                   IsNull(Cta.MovAgo,0) MovAgo," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P308' THEN p.nMonIni ELSE 0 END)" & _
                "                   - IsNull(Cta.MovAgo,0)) DMAgo, "
            tmpSql = tmpSql & _
                "                   sum(CASE WHEN p.cPeriodo = 'P309' THEN p.nMonIni ELSE 0 END) Seti," & _
                "                   IsNull(Cta.MovSet,0) MovSet," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P309' THEN p.nMonIni ELSE 0 END)" & _
                "                   - IsNull(Cta.MovSet,0)) DMSet, " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P310' THEN p.nMonIni ELSE 0 END) Oct," & _
                "                   IsNull(Cta.MovOct,0) MovOct," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P310' THEN p.nMonIni ELSE 0 END)" & _
                "                   - IsNull(Cta.MovOct,0)) DMOct, " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P311' THEN p.nMonIni ELSE 0 END) Nov," & _
                "                   IsNull(Cta.MovNov,0) MovNov," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P311' THEN p.nMonIni ELSE 0 END)" & _
                "                   - IsNull(Cta.MovNov,0)) DMNov, " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P312' THEN p.nMonIni ELSE 0 END) Dic," & _
                "                   IsNull(Cta.MovDic,0) MovDic," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P312' THEN p.nMonIni ELSE 0 END)" & _
                "                   - IsNull(Cta.MovDic,0)) DMDic "
            tmpSql = tmpSql & " FROM PRubro R LEFT JOIN pPresupu P ON r.cAno = p.cAno AND p.cPresu = r.cPresu AND p.cCodRub = r.cCodRub" & _
                "                   LEFT JOIN " & _
                "                       (SELECT rc.cAno, rc.cPresu, rc.cCodRub, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____01%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovEne, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____02%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovFeb, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____03%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovMar, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____04%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovAbr, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____05%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovMay, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____06%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovJun, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____07%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovJul, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____08%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovAgo, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____09%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovSet, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____10%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovOct, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____11%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovNov, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____12%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovDic " & _
                "                       FROM PRubCta RC JOIN MovCta MC ON SubString(mc.cCtaContCod,1,2) + '_' + SubString(mc.cCtaContCod,4,len(mc.cCtaContCod)) like SubString(rc.cCtaCnt,1,2) + '_' + SubString(rc.cCtaCnt,4,len(rc.cCtaCnt)) + '%' " & _
                "                           JOIN MOV M ON mc.cMovNro = m.cMovNro "
            tmpSql = tmpSql & _
                "                           JOIN (SELECT cValor, cAbrev FROM " & gcCentralCom & "TablaCod " & _
                "                                   WHERE cCodTab LIKE 'C0__') Tip " & _
                "                           ON substring(mc.cCtaContCod,1,len(tip.cValor)) = Tip.cValor "
            tmpSql = tmpSql & _
                "                       WHERE mc.cMovNro LIKE '" & Left(Trim(cboFecha.Text), 4) & "%' " & _
                "                           AND SubString(mc.cCtaContCod,3,1) IN ('1','2') " & _
                "                           AND m.cMovFlag <> 'X' AND (m.cMovEstado='9' OR " & _
                "                                   (m.cMovEstado='0' AND Not Exists(Select Distinct m2.cmovestado " & _
                "                                                           From mov m2 join movref mr " & _
                "                                                               On m2.cmovnro = mr.cmovnroref " & _
                "                                                               And mr.cmovnro = mc.cmovnro " & _
                "                                                           Where m2.cMovFlag <> 'X' And m2.cMovEstado = '9'))) " & _
                "                       GROUP BY rc.cAno, rc.cPresu, rc.cCodRub, tip.cAbrev ) Cta " & _
                "                   ON r.cAno = cta.cAno AND r.cPresu = cta.cPresu AND r.cCodRub = cta.cCodRub " & _
                "               WHERE r.cAno = '" & Left(Trim(cboFecha.Text), 4) & "' AND r.cPresu = '" & Right(cboPresu.Text, 4) & "'" & _
                "                   AND r.cFlag Is Null" & _
                "               GROUP BY r.cCodRub, r.cDesRub, r.nMonIni, r.nMonto, r.nMonCre1, r.nMonCre2, " & _
                "                   Cta.MovEne, Cta.MovFeb, Cta.MovMar, Cta.MovAbr, Cta.MovMay, Cta.MovJun, " & _
                "                   Cta.MovJul, Cta.MovAgo, Cta.MovSet, Cta.MovOct, Cta.MovNov, Cta.MovDic "
        End If
        tmpSql = tmpSql + " )Res" & _
            "           ON SUBSTRING(res.cCodRub,1,LEN(rr.cCodRub)) = rr.cCodRub " & _
            " GROUP BY rr.cCodRub, rr.cDesRub " & _
            " ORDER BY rr.cCodRub "
    ElseIf pnTipo = 2 Then   'Trimestral
        tmpSql = "SELECT rr.cCodRub, rr.cDesRub, sum(res.nMonIni) nMonIni, sum(res.nMonto) nMonto, sum(res.nMonCre1) nMonCre1, sum(res.nMonCre2) nMonCre2, " & _
            "       SUM(res.nMonto + res.nMonCre1 + res.nMonCre2) Total, " & _
            "       SUM(Primero) Primero, SUM(MovPrimero) MovPrimero, SUM(DMPrimero) DMPrimero, " & _
            "       SUM(Segundo) Segundo, SUM(MovSegundo) MovSegundo, SUM(DMSegundo) DMSegundo, " & _
            "       SUM(Tercero) Tercero, SUM(MovTercero) MovTercero, SUM(DMTercero) DMTercero, " & _
            "       SUM(Cuarto) Cuarto, SUM(MovCuarto) MovCuarto, SUM(DMCuarto) DMCuarto " & _
            " FROM (SELECT DISTINCT cCodRub, cDesRub " & _
            "       FROM pRubro " & _
            "       WHERE cAno = '" & Left(Trim(cboFecha.Text), 4) & "' AND cPresu = '" & Right(cboPresu.Text, 4) & "' " & _
            "       ) RR " & _
            "           LEFT JOIN " & _
            "              (    "
        If chkProy.value = 1 Then
            'Movimientos CON Proyeccion
            tmpSql = tmpSql + " SELECT r.cCodRub, r.cDesRub, r.nMonIni, r.nMonto, r.nMonCre1, r.nMonCre2, " & _
                "                   (r.nMonto + r.nMonCre1 + r.nMonCre2) Total ," & _
                "                   sum(CASE WHEN p.cPeriodo = 'P301' OR p.cPeriodo = 'P302' OR p.cPeriodo = 'P303' THEN p.nMonIni ELSE 0 END) Primero," & _
                "                   (IsNull(Cta.MovEne,0) + IsNull(Cta.MovFeb,0) + IsNull(Cta.MovMar, 0) + IsNull(Pro.ProEne,0) + IsNull(Pro.ProFeb,0) + IsNull(Pro.ProMar,0)) MovPrimero, " & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P301' OR p.cPeriodo = 'P302' OR p.cPeriodo = 'P303' THEN p.nMonIni ELSE 0 END)" & _
                "                   - (IsNull(Cta.MovEne,0) + IsNull(Cta.MovFeb,0) + IsNull(Cta.MovMar, 0) + IsNull(Pro.ProEne,0) + IsNull(Pro.ProFeb,0) + IsNull(Pro.ProMar,0))) DMPrimero, " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P304' OR p.cPeriodo = 'P305' OR p.cPeriodo = 'P306' THEN p.nMonIni ELSE 0 END) Segundo," & _
                "                   (IsNull(Cta.MovAbr,0) + IsNull(Cta.MovMay,0) + IsNull(Cta.MovJun, 0) + IsNull(Pro.ProAbr,0) + IsNull(Pro.ProMay,0) + IsNull(Pro.ProJun, 0)) MovSegundo," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P304' OR p.cPeriodo = 'P305' OR p.cPeriodo = 'P306' THEN p.nMonIni ELSE 0 END)" & _
                "                   - (IsNull(Cta.MovAbr,0) + IsNull(Cta.MovMay,0) + IsNull(Cta.MovJun, 0) + IsNull(Pro.ProAbr,0) + IsNull(Pro.ProMay,0) + IsNull(Pro.ProJun, 0))) DMSegundo, " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P307' OR p.cPeriodo = 'P308' OR p.cPeriodo = 'P309' THEN p.nMonIni ELSE 0 END) Tercero," & _
                "                   (IsNull(Cta.MovJul,0) + IsNull(Cta.MovAgo,0) + IsNull(Cta.MovSet, 0) + IsNull(Pro.ProJul,0) + IsNull(Pro.ProAgo,0) + IsNull(Pro.ProSet,0)) MovTercero," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P307' OR p.cPeriodo = 'P308' OR p.cPeriodo = 'P309' THEN p.nMonIni ELSE 0 END)" & _
                "                   - (IsNull(Cta.MovJul,0) + IsNull(Cta.MovAgo,0) + IsNull(Cta.MovSet, 0) + IsNull(Pro.ProJul,0) + IsNull(Pro.ProAgo,0) + IsNull(Pro.ProSet,0))) DMTercero,  " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P310' OR p.cPeriodo = 'P311' OR p.cPeriodo = 'P312' THEN p.nMonIni ELSE 0 END) Cuarto," & _
                "                   (IsNull(Cta.MovOct,0) + IsNull(Cta.MovNov,0) + IsNull(Cta.MovDic, 0) + IsNull(Pro.ProOct,0) + IsNull(Pro.ProNov,0) + IsNull(Pro.ProDic, 0)) MovCuarto," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P310' OR p.cPeriodo = 'P311' OR p.cPeriodo = 'P312' THEN p.nMonIni ELSE 0 END)" & _
                "                   - (IsNull(Cta.MovOct,0) + IsNull(Cta.MovNov,0) + IsNull(Cta.MovDic, 0) + IsNull(Pro.ProOct,0) + IsNull(Pro.ProNov,0) + IsNull(Pro.ProDic, 0))) DMCuarto "
            tmpSql = tmpSql & " FROM PRubro R LEFT JOIN pPresupu P ON r.cAno = p.cAno AND p.cPresu = r.cPresu AND p.cCodRub = r.cCodRub" & _
                "                   LEFT JOIN " & _
                "                       (SELECT rc.cAno, rc.cPresu, rc.cCodRub, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____01%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovEne, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____02%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovFeb, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____03%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovMar, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____04%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovAbr, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____05%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovMay, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____06%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovJun, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____07%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovJul, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____08%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovAgo, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____09%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovSet, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____10%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovOct, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____11%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovNov, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____12%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovDic " & _
                "                       FROM PRubCta RC JOIN MovCta MC ON SubString(mc.cCtaContCod,1,2) + '_' + SubString(mc.cCtaContCod,4,len(mc.cCtaContCod)) like SubString(rc.cCtaCnt,1,2) + '_' + SubString(rc.cCtaCnt,4,len(rc.cCtaCnt)) + '%' " & _
                "                           JOIN MOV M ON mc.cMovNro = m.cMovNro "
            tmpSql = tmpSql & _
                "                           JOIN (SELECT cValor, cAbrev FROM " & gcCentralCom & "TablaCod " & _
                "                                   WHERE cCodTab LIKE 'C0__') Tip " & _
                "                           ON substring(mc.cCtaContCod,1,len(tip.cValor)) = Tip.cValor "
            tmpSql = tmpSql & _
                "                       WHERE mc.cMovNro LIKE '" & Left(Trim(cboFecha.Text), 4) & "%' " & _
                "                           AND SubString(mc.cCtaContCod,3,1) IN ('1','2') " & _
                "                           AND m.cMovFlag <> 'X' AND (m.cMovEstado='9' OR " & _
                "                                   (m.cMovEstado='0' AND Not Exists(Select Distinct m2.cmovestado " & _
                "                                                           From mov m2 join movref mr " & _
                "                                                               On m2.cmovnro = mr.cmovnroref " & _
                "                                                               And mr.cmovnro = mc.cmovnro " & _
                "                                                           Where m2.cMovFlag <> 'X' And m2.cMovEstado = '9'))) " & _
                "                       GROUP BY rc.cAno, rc.cPresu, rc.cCodRub, tip.cAbrev ) Cta " & _
                "                   ON r.cAno = cta.cAno AND r.cPresu = cta.cPresu AND r.cCodRub = cta.cCodRub "
            tmpSql = tmpSql & _
                "                   LEFT JOIN " & _
                "                       (SELECT  rc.cAno, rc.cPresu, rc.cCodRub, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 1 THEN (msd.nServImporte) ELSE 0 END),0)) ProEne, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 2 THEN (msd.nServImporte) ELSE 0 END),0)) ProFeb, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 3 THEN (msd.nServImporte) ELSE 0 END),0)) ProMar, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 4 THEN (msd.nServImporte) ELSE 0 END),0)) ProAbr, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 5 THEN (msd.nServImporte) ELSE 0 END),0)) ProMay, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 6 THEN (msd.nServImporte) ELSE 0 END),0)) ProJun, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 7 THEN (msd.nServImporte) ELSE 0 END),0)) ProJul, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 8 THEN (msd.nServImporte) ELSE 0 END),0)) ProAgo, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 9 THEN (msd.nServImporte) ELSE 0 END),0)) ProSet, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 10 THEN (msd.nServImporte) ELSE 0 END),0)) ProOct, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 11 THEN (msd.nServImporte) ELSE 0 END),0)) ProNov, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 12 THEN (msd.nServImporte) ELSE 0 END),0)) ProDic " & _
                "                       FROM MovServicioDet MSD JOIN MovServicio MS ON msd.cMovNro = ms.cMovNro " & _
                "                           JOIN Mov M ON ms.cMovNro = m.cMovNro JOIN MovCta MC ON m.cMovNro = mc.cMovNro " & _
                "                           JOIN PRubCta RC ON mc.cCtaContCod = rc.cCtaCnt  AND rc.cPresu = ms.cPresuCod " & _
                "                       WHERE  Year(dServFecha) = '" & Left(Trim(cboFecha.Text), 4) & "' AND month(dServFecha) >= " & Month(gdFecSis) & " " & _
                "                           AND m.cMovEstado='S' AND m.cMovFlag <> 'X' " & _
                "                       GROUP BY rc.cAno, rc.cPresu, rc.cCodRub ) Pro " & _
                "                   ON r.cAno = pro.cAno AND r.cPresu = pro.cPresu AND r.cCodRub = pro.cCodRub "
            tmpSql = tmpSql & _
                "               WHERE r.cAno = '" & Left(Trim(cboFecha.Text), 4) & "' AND r.cPresu = '" & Right(cboPresu.Text, 4) & "'" & _
                "                   AND r.cFlag Is Null" & _
                "               GROUP BY r.cCodRub, r.cDesRub, r.nMonIni, r.nMonto, r.nMonCre1, r.nMonCre2, " & _
                "                   Cta.MovEne, Cta.MovFeb, Cta.MovMar, Cta.MovAbr, Cta.MovMay, Cta.MovJun, " & _
                "                   Cta.MovJul, Cta.MovAgo, Cta.MovSet, Cta.MovOct, Cta.MovNov, Cta.MovDic, " & _
                "                   Pro.ProEne, Pro.ProFeb, Pro.ProMar, Pro.ProAbr, Pro.ProMay, Pro.ProJun, " & _
                "                   Pro.ProJul, Pro.ProAgo, Pro.ProSet, Pro.ProOct, Pro.ProNov, Pro.ProDic "
        Else
            'Movimientos SIN Proyección
            tmpSql = tmpSql + " SELECT r.cCodRub, r.cDesRub, r.nMonIni, r.nMonto, r.nMonCre1, r.nMonCre2, " & _
                "                   (r.nMonto + r.nMonCre1 + r.nMonCre2) Total ," & _
                "                   sum(CASE WHEN p.cPeriodo = 'P301' OR p.cPeriodo = 'P302' OR p.cPeriodo = 'P303' THEN p.nMonIni ELSE 0 END) Primero," & _
                "                   (IsNull(Cta.MovEne,0) + IsNull(Cta.MovFeb,0) + IsNull(Cta.MovMar, 0)) MovPrimero, " & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P301' OR p.cPeriodo = 'P302' OR p.cPeriodo = 'P303' THEN p.nMonIni ELSE 0 END)" & _
                "                   - (IsNull(Cta.MovEne,0) + IsNull(Cta.MovFeb,0) + IsNull(Cta.MovMar, 0))) DMPrimero, " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P304' OR p.cPeriodo = 'P305' OR p.cPeriodo = 'P306' THEN p.nMonIni ELSE 0 END) Segundo," & _
                "                   (IsNull(Cta.MovAbr,0) + IsNull(Cta.MovMay,0) + IsNull(Cta.MovJun, 0)) MovSegundo," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P304' OR p.cPeriodo = 'P305' OR p.cPeriodo = 'P306' THEN p.nMonIni ELSE 0 END)" & _
                "                   - (IsNull(Cta.MovAbr,0) + IsNull(Cta.MovMay,0) + IsNull(Cta.MovJun, 0))) DMSegundo, " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P307' OR p.cPeriodo = 'P308' OR p.cPeriodo = 'P309' THEN p.nMonIni ELSE 0 END) Tercero," & _
                "                   (IsNull(Cta.MovJul,0) + IsNull(Cta.MovAgo,0) + IsNull(Cta.MovSet, 0)) MovTercero," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P307' OR p.cPeriodo = 'P308' OR p.cPeriodo = 'P309' THEN p.nMonIni ELSE 0 END)" & _
                "                   - (IsNull(Cta.MovJul,0) + IsNull(Cta.MovAgo,0) + IsNull(Cta.MovSet, 0))) DMTercero,  " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P310' OR p.cPeriodo = 'P311' OR p.cPeriodo = 'P312' THEN p.nMonIni ELSE 0 END) Cuarto," & _
                "                   (IsNull(Cta.MovOct,0) + IsNull(Cta.MovNov,0) + IsNull(Cta.MovDic, 0)) MovCuarto," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P310' OR p.cPeriodo = 'P311' OR p.cPeriodo = 'P312' THEN p.nMonIni ELSE 0 END)" & _
                "                   - (IsNull(Cta.MovOct,0) + IsNull(Cta.MovNov,0) + IsNull(Cta.MovDic, 0))) DMCuarto "
            tmpSql = tmpSql & " FROM PRubro R LEFT JOIN pPresupu P ON r.cAno = p.cAno AND p.cPresu = r.cPresu AND p.cCodRub = r.cCodRub" & _
                "                   LEFT JOIN " & _
                "                       (SELECT rc.cAno, rc.cPresu, rc.cCodRub, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____01%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovEne, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____02%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovFeb, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____03%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovMar, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____04%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovAbr, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____05%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovMay, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____06%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovJun, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____07%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovJul, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____08%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovAgo, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____09%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovSet, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____10%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovOct, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____11%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovNov, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____12%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovDic " & _
                "                       FROM PRubCta RC JOIN MovCta MC ON SubString(mc.cCtaContCod,1,2) + '_' + SubString(mc.cCtaContCod,4,len(mc.cCtaContCod)) like SubString(rc.cCtaCnt,1,2) + '_' + SubString(rc.cCtaCnt,4,len(rc.cCtaCnt)) + '%' " & _
                "                           JOIN MOV M ON mc.cMovNro = m.cMovNro "
            tmpSql = tmpSql & _
                "                           JOIN (SELECT cValor, cAbrev FROM " & gcCentralCom & "TablaCod " & _
                "                                   WHERE cCodTab LIKE 'C0__') Tip " & _
                "                           ON substring(mc.cCtaContCod,1,len(tip.cValor)) = Tip.cValor "
            tmpSql = tmpSql & _
                "                       WHERE mc.cMovNro LIKE '" & Left(Trim(cboFecha.Text), 4) & "%' " & _
                "                           AND SubString(mc.cCtaContCod,3,1) IN ('1','2') " & _
                "                           AND m.cMovFlag <> 'X' AND (m.cMovEstado='9' OR " & _
                "                                   (m.cMovEstado='0' AND Not Exists(Select Distinct m2.cmovestado " & _
                "                                                           From mov m2 join movref mr " & _
                "                                                               On m2.cmovnro = mr.cmovnroref " & _
                "                                                               And mr.cmovnro = mc.cmovnro " & _
                "                                                           Where m2.cMovFlag <> 'X' And m2.cMovEstado = '9'))) " & _
                "                       GROUP BY rc.cAno, rc.cPresu, rc.cCodRub, tip.cAbrev ) Cta " & _
                "                   ON r.cAno = cta.cAno AND r.cPresu = cta.cPresu AND r.cCodRub = cta.cCodRub " & _
                "               WHERE r.cAno = '" & Left(Trim(cboFecha.Text), 4) & "' AND r.cPresu = '" & Right(cboPresu.Text, 4) & "'" & _
                "                   AND r.cFlag Is Null" & _
                "               GROUP BY r.cCodRub, r.cDesRub, r.nMonIni, r.nMonto, r.nMonCre1, r.nMonCre2, " & _
                "                   Cta.MovEne, Cta.MovFeb, Cta.MovMar, Cta.MovAbr, Cta.MovMay, Cta.MovJun, " & _
                "                   Cta.MovJul, Cta.MovAgo, Cta.MovSet, Cta.MovOct, Cta.MovNov, Cta.MovDic "
        End If
        tmpSql = tmpSql + " )Res" & _
            "           ON SUBSTRING(res.cCodRub,1,LEN(rr.cCodRub)) = rr.cCodRub " & _
            " GROUP BY rr.cCodRub, rr.cDesRub " & _
            " ORDER BY rr.cCodRub "
    Else   'SEMESTRAL
        'nTipo = 3
        tmpSql = "SELECT rr.cCodRub, rr.cDesRub, sum(res.nMonIni) nMonIni, sum(res.nMonto) nMonto, sum(res.nMonCre1) nMonCre1, sum(res.nMonCre2) nMonCre2, " & _
            "       SUM(res.nMonto + res.nMonCre1 + res.nMonCre2) Total, " & _
            "       SUM(Primero) Primero, SUM(MovPrimero) MovPrimero, SUM(DMPrimero) DMPrimero, " & _
            "       SUM(Segundo) Segundo, SUM(MovSegundo) MovSegundo, SUM(DMSegundo) DMSegundo " & _
            " FROM (SELECT DISTINCT cCodRub, cDesRub " & _
            "       FROM pRubro " & _
            "       WHERE cAno = '" & Left(Trim(cboFecha.Text), 4) & "' AND cPresu = '" & Right(cboPresu.Text, 4) & "' " & _
            "       ) RR " & _
            "           LEFT JOIN " & _
            "              (    "
        If chkProy.value = 1 Then
            'Movimientos CON Proyeccion
            tmpSql = tmpSql & " SELECT r.cCodRub, r.cDesRub, r.nMonIni, r.nMonto, r.nMonCre1, r.nMonCre2, " & _
                "                   (r.nMonto + r.nMonCre1 + r.nMonCre2) Total ," & _
                "                   sum(CASE WHEN p.cPeriodo = 'P301' OR p.cPeriodo = 'P302' OR p.cPeriodo = 'P303' OR p.cPeriodo = 'P304' OR p.cPeriodo = 'P305' OR p.cPeriodo = 'P306' THEN p.nMonIni ELSE 0 END) Primero," & _
                "                   (IsNull(Cta.MovEne,0) + IsNull(Cta.MovFeb,0) + IsNull(Cta.MovMar,0) + IsNull(Cta.MovAbr,0) + IsNull(Cta.MovMay,0) + IsNull(Cta.MovJun, 0) + IsNull(Pro.ProEne,0) + IsNull(Pro.ProFeb,0) + IsNull(Pro.ProMar,0) + IsNull(Pro.ProAbr,0) + IsNull(Pro.ProMay,0) + IsNull(Pro.ProJun, 0)) MovPrimero, " & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P301' OR p.cPeriodo = 'P302' OR p.cPeriodo = 'P303' OR p.cPeriodo = 'P304' OR p.cPeriodo = 'P305' OR p.cPeriodo = 'P306' THEN p.nMonIni ELSE 0 END)" & _
                "                   - (IsNull(Cta.MovEne,0) + IsNull(Cta.MovFeb,0) + IsNull(Cta.MovMar,0) + IsNull(Cta.MovAbr,0) + IsNull(Cta.MovMay,0) + IsNull(Cta.MovJun, 0) + IsNull(Pro.ProEne,0) + IsNull(Pro.ProFeb,0) + IsNull(Pro.ProMar,0) + IsNull(Pro.ProAbr,0) + IsNull(Pro.ProMay,0) + IsNull(Pro.ProJun, 0))) DMPrimero, " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P307' OR p.cPeriodo = 'P308' OR p.cPeriodo = 'P309' OR p.cPeriodo = 'P310' OR p.cPeriodo = 'P311' OR p.cPeriodo = 'P312' THEN p.nMonIni ELSE 0 END) Segundo," & _
                "                   (IsNull(Cta.MovJul,0) + IsNull(Cta.MovAgo,0) + IsNull(Cta.MovSet,0) + IsNull(Cta.MovOct,0) + IsNull(Cta.MovNov,0) + IsNull(Cta.MovDic, 0) + IsNull(Pro.ProJul,0) + IsNull(Pro.ProAgo,0) + IsNull(Pro.ProSet,0) + IsNull(Pro.ProOct,0) + IsNull(Pro.ProNov,0) + IsNull(Pro.ProDic, 0)) MovSegundo," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P307' OR p.cPeriodo = 'P308' OR p.cPeriodo = 'P309' OR p.cPeriodo = 'P310' OR p.cPeriodo = 'P311' OR p.cPeriodo = 'P312' THEN p.nMonIni ELSE 0 END)" & _
                "                   - (IsNull(Cta.MovJul,0) + IsNull(Cta.MovAgo,0) + IsNull(Cta.MovSet,0) + IsNull(Cta.MovOct,0) + IsNull(Cta.MovNov,0) + IsNull(Cta.MovDic, 0) + IsNull(Pro.ProJul,0) + IsNull(Pro.ProAgo,0) + IsNull(Pro.ProSet,0) + IsNull(Pro.ProOct,0) + IsNull(Pro.ProNov,0) + IsNull(Pro.ProDic, 0))) DMSegundo  "
            tmpSql = tmpSql & " FROM PRubro R LEFT JOIN pPresupu P ON r.cAno = p.cAno AND p.cPresu = r.cPresu AND p.cCodRub = r.cCodRub" & _
                "                   LEFT JOIN " & _
                "                       (SELECT rc.cAno, rc.cPresu, rc.cCodRub, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____01%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovEne, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____02%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovFeb, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____03%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovMar, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____04%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovAbr, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____05%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovMay, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____06%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovJun, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____07%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovJul, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____08%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovAgo, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____09%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovSet, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____10%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovOct, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____11%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovNov, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____12%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovDic " & _
                "                       FROM PRubCta RC JOIN MovCta MC ON SubString(mc.cCtaContCod,1,2) + '_' + SubString(mc.cCtaContCod,4,len(mc.cCtaContCod)) like SubString(rc.cCtaCnt,1,2) + '_' + SubString(rc.cCtaCnt,4,len(rc.cCtaCnt)) + '%' " & _
                "                           JOIN MOV M ON mc.cMovNro = m.cMovNro "
            tmpSql = tmpSql & _
                "                           JOIN (SELECT cValor, cAbrev FROM " & gcCentralCom & "TablaCod " & _
                "                                   WHERE cCodTab LIKE 'C0__') Tip " & _
                "                           ON substring(mc.cCtaContCod,1,len(tip.cValor)) = Tip.cValor "
            tmpSql = tmpSql & _
                "                       WHERE mc.cMovNro LIKE '" & Left(Trim(cboFecha.Text), 4) & "%' " & _
                "                           AND SubString(mc.cCtaContCod,3,1) IN ('1','2') " & _
                "                           AND m.cMovFlag <> 'X' AND (m.cMovEstado='9' OR " & _
                "                                   (m.cMovEstado='0' AND Not Exists(Select Distinct m2.cmovestado " & _
                "                                                           From mov m2 join movref mr " & _
                "                                                               On m2.cmovnro = mr.cmovnroref " & _
                "                                                               And mr.cmovnro = mc.cmovnro " & _
                "                                                           Where m2.cMovFlag <> 'X' And m2.cMovEstado = '9'))) " & _
                "                       GROUP BY rc.cAno, rc.cPresu, rc.cCodRub, tip.cAbrev ) Cta " & _
                "                   ON r.cAno = cta.cAno AND r.cPresu = cta.cPresu AND r.cCodRub = cta.cCodRub "
            tmpSql = tmpSql & _
                "                   LEFT JOIN " & _
                "                       (SELECT  rc.cAno, rc.cPresu, rc.cCodRub, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 1 THEN (msd.nServImporte) ELSE 0 END),0)) ProEne, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 2 THEN (msd.nServImporte) ELSE 0 END),0)) ProFeb, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 3 THEN (msd.nServImporte) ELSE 0 END),0)) ProMar, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 4 THEN (msd.nServImporte) ELSE 0 END),0)) ProAbr, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 5 THEN (msd.nServImporte) ELSE 0 END),0)) ProMay, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 6 THEN (msd.nServImporte) ELSE 0 END),0)) ProJun, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 7 THEN (msd.nServImporte) ELSE 0 END),0)) ProJul, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 8 THEN (msd.nServImporte) ELSE 0 END),0)) ProAgo, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 9 THEN (msd.nServImporte) ELSE 0 END),0)) ProSet, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 10 THEN (msd.nServImporte) ELSE 0 END),0)) ProOct, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 11 THEN (msd.nServImporte) ELSE 0 END),0)) ProNov, " & _
                "                           Abs(IsNull(Sum(CASE WHEN month(dServFecha) = 12 THEN (msd.nServImporte) ELSE 0 END),0)) ProDic " & _
                "                       FROM MovServicioDet MSD JOIN MovServicio MS ON msd.cMovNro = ms.cMovNro " & _
                "                           JOIN Mov M ON ms.cMovNro = m.cMovNro JOIN MovCta MC ON m.cMovNro = mc.cMovNro " & _
                "                           JOIN PRubCta RC ON mc.cCtaContCod = rc.cCtaCnt  AND rc.cPresu = ms.cPresuCod " & _
                "                       WHERE  Year(dServFecha) = '" & Left(Trim(cboFecha.Text), 4) & "' AND month(dServFecha) >= " & Month(gdFecSis) & " " & _
                "                           AND m.cMovEstado='S' AND m.cMovFlag <> 'X' " & _
                "                       GROUP BY rc.cAno, rc.cPresu, rc.cCodRub ) Pro " & _
                "                   ON r.cAno = pro.cAno AND r.cPresu = pro.cPresu AND r.cCodRub = pro.cCodRub "
            tmpSql = tmpSql & _
                "               WHERE r.cAno = '" & Left(Trim(cboFecha.Text), 4) & "' AND r.cPresu = '" & Right(cboPresu.Text, 4) & "'" & _
                "                   AND r.cFlag Is Null" & _
                "               GROUP BY r.cCodRub, r.cDesRub, r.nMonIni, r.nMonto, r.nMonCre1, r.nMonCre2, " & _
                "                   Cta.MovEne, Cta.MovFeb, Cta.MovMar, Cta.MovAbr, Cta.MovMay, Cta.MovJun, " & _
                "                   Cta.MovJul, Cta.MovAgo, Cta.MovSet, Cta.MovOct, Cta.MovNov, Cta.MovDic, " & _
                "                   Pro.ProEne, Pro.ProFeb, Pro.ProMar, Pro.ProAbr, Pro.ProMay, Pro.ProJun, " & _
                "                   Pro.ProJul, Pro.ProAgo, Pro.ProSet, Pro.ProOct, Pro.ProNov, Pro.ProDic "
        Else
            'Movimientos SIN Proyección
            tmpSql = tmpSql & " SELECT r.cCodRub, r.cDesRub, r.nMonIni, r.nMonto, r.nMonCre1, r.nMonCre2, " & _
                "                   (r.nMonto + r.nMonCre1 + r.nMonCre2) Total ," & _
                "                   sum(CASE WHEN p.cPeriodo = 'P301' OR p.cPeriodo = 'P302' OR p.cPeriodo = 'P303' OR p.cPeriodo = 'P304' OR p.cPeriodo = 'P305' OR p.cPeriodo = 'P306' THEN p.nMonIni ELSE 0 END) Primero," & _
                "                   (IsNull(Cta.MovEne,0) + IsNull(Cta.MovFeb,0) + IsNull(Cta.MovMar,0) + IsNull(Cta.MovAbr,0) + IsNull(Cta.MovMay,0) + IsNull(Cta.MovJun, 0)) MovPrimero, " & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P301' OR p.cPeriodo = 'P302' OR p.cPeriodo = 'P303' OR p.cPeriodo = 'P304' OR p.cPeriodo = 'P305' OR p.cPeriodo = 'P306' THEN p.nMonIni ELSE 0 END)" & _
                "                   - (IsNull(Cta.MovEne,0) + IsNull(Cta.MovFeb,0) + IsNull(Cta.MovMar,0) + IsNull(Cta.MovAbr,0) + IsNull(Cta.MovMay,0) + IsNull(Cta.MovJun, 0))) DMPrimero, " & _
                "                   sum(CASE WHEN p.cPeriodo = 'P307' OR p.cPeriodo = 'P308' OR p.cPeriodo = 'P309' OR p.cPeriodo = 'P310' OR p.cPeriodo = 'P311' OR p.cPeriodo = 'P312' THEN p.nMonIni ELSE 0 END) Segundo," & _
                "                   (IsNull(Cta.MovJul,0) + IsNull(Cta.MovAgo,0) + IsNull(Cta.MovSet,0) + IsNull(Cta.MovOct,0) + IsNull(Cta.MovNov,0) + IsNull(Cta.MovDic, 0)) MovSegundo," & _
                "                   (sum(CASE WHEN p.cPeriodo = 'P307' OR p.cPeriodo = 'P308' OR p.cPeriodo = 'P309' OR p.cPeriodo = 'P310' OR p.cPeriodo = 'P311' OR p.cPeriodo = 'P312' THEN p.nMonIni ELSE 0 END)" & _
                "                   - (IsNull(Cta.MovJul,0) + IsNull(Cta.MovAgo,0) + IsNull(Cta.MovSet,0) + IsNull(Cta.MovOct,0) + IsNull(Cta.MovNov,0) + IsNull(Cta.MovDic, 0))) DMSegundo  "
            tmpSql = tmpSql & " FROM PRubro R LEFT JOIN pPresupu P ON r.cAno = p.cAno AND p.cPresu = r.cPresu AND p.cCodRub = r.cCodRub" & _
                "                   LEFT JOIN " & _
                "                       (SELECT rc.cAno, rc.cPresu, rc.cCodRub, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____01%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovEne, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____02%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovFeb, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____03%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovMar, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____04%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovAbr, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____05%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovMay, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____06%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovJun, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____07%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovJul, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____08%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovAgo, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____09%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovSet, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____10%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovOct, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____11%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovNov, " & _
                "                           (IsNull(Sum(CASE WHEN mc.cMovNro like '____12%' THEN (CASE WHEN tip.cAbrev = 'D' THEN (mc.nMovImporte) ELSE (mc.nMovImporte * -1) END) ELSE 0 END),0)) MovDic " & _
                "                       FROM PRubCta RC JOIN MovCta MC ON SubString(mc.cCtaContCod,1,2) + '_' + SubString(mc.cCtaContCod,4,len(mc.cCtaContCod)) like SubString(rc.cCtaCnt,1,2) + '_' + SubString(rc.cCtaCnt,4,len(rc.cCtaCnt)) + '%' " & _
                "                           JOIN MOV M ON mc.cMovNro = m.cMovNro "
            tmpSql = tmpSql & _
                "                           JOIN (SELECT cValor, cAbrev FROM " & gcCentralCom & "TablaCod " & _
                "                                   WHERE cCodTab LIKE 'C0__') Tip " & _
                "                           ON substring(mc.cCtaContCod,1,len(tip.cValor)) = Tip.cValor "
            tmpSql = tmpSql & _
                "                       WHERE mc.cMovNro LIKE '" & Left(Trim(cboFecha.Text), 4) & "%' " & _
                "                           AND SubString(mc.cCtaContCod,3,1) IN ('1','2') " & _
                "                           AND m.cMovFlag <> 'X' AND (m.cMovEstado='9' OR " & _
                "                                   (m.cMovEstado='0' AND Not Exists(Select Distinct m2.cmovestado " & _
                "                                                           From mov m2 join movref mr " & _
                "                                                               On m2.cmovnro = mr.cmovnroref " & _
                "                                                               And mr.cmovnro = mc.cmovnro " & _
                "                                                           Where m2.cMovFlag <> 'X' And m2.cMovEstado = '9'))) " & _
                "                       GROUP BY rc.cAno, rc.cPresu, rc.cCodRub, tip.cAbrev ) Cta " & _
                "                   ON r.cAno = cta.cAno AND r.cPresu = cta.cPresu AND r.cCodRub = cta.cCodRub " & _
                "               WHERE r.cAno = '" & Left(Trim(cboFecha.Text), 4) & "' AND r.cPresu = '" & Right(cboPresu.Text, 4) & "'" & _
                "                   AND r.cFlag Is Null" & _
                "               GROUP BY r.cCodRub, r.cDesRub, r.nMonIni, r.nMonto, r.nMonCre1, r.nMonCre2, " & _
                "                   Cta.MovEne, Cta.MovFeb, Cta.MovMar, Cta.MovAbr, Cta.MovMay, Cta.MovJun, " & _
                "                   Cta.MovJul, Cta.MovAgo, Cta.MovSet, Cta.MovOct, Cta.MovNov, Cta.MovDic "
        End If
        tmpSql = tmpSql + " )Res" & _
            "           ON SUBSTRING(res.cCodRub,1,LEN(rr.cCodRub)) = rr.cCodRub " & _
            " GROUP BY rr.cCodRub, rr.cDesRub " & _
            " ORDER BY rr.cCodRub "
    End If
    
    CargaSelect = tmpSql
End Function

Private Sub CuadroExcel(plHoja1 As Excel.Worksheet, X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer, Optional lbLineasVert As Boolean = False)
Dim i, j As Integer

For i = X1 To X2
    plHoja1.Range(plHoja1.Cells(Y1, i), plHoja1.Cells(Y1, i)).Borders(xlEdgeTop).LineStyle = xlContinuous
    plHoja1.Range(plHoja1.Cells(Y2, i), plHoja1.Cells(Y2, i)).Borders(xlEdgeBottom).LineStyle = xlContinuous
Next i
If lbLineasVert = False Then
    For i = X1 To X2
        For j = Y1 To Y2
            plHoja1.Range(plHoja1.Cells(j, i), plHoja1.Cells(j, i)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Next j
    Next i
End If
If lbLineasVert Then
    For j = Y1 To Y2
        plHoja1.Range(plHoja1.Cells(j, X1), plHoja1.Cells(j, X1)).Borders(xlEdgeRight).LineStyle = xlContinuous
    Next j
End If

For j = Y1 To Y2
    plHoja1.Range(plHoja1.Cells(j, X2), plHoja1.Cells(j, X2)).Borders(xlEdgeRight).LineStyle = xlContinuous
Next j
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Dim tmpSql As String
    Dim oPP As DPresupuesto
    Set oPP = New DPresupuesto
    Dim oCon As DConstantes
    Set oCon = New DConstantes
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim clsDGnral As DLogGeneral
    Set clsDGnral = New DLogGeneral
    
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
    
    cboMonedaP.Clear
    cboMonedaP.AddItem "NACIONAL                                 1"
    cboMonedaP.AddItem "EXTRANJERA                               2"
    cboMonedaP.AddItem "CONSOL (MN ME)                           3"
    cboMonedaP.AddItem "AJUSTE                                   4"
    cboMonedaP.AddItem "NACIONAL CON AJUSTE                      5"
    cboMonedaP.AddItem "EXTRANJERA CON AJUSTE                    6"
    cboMonedaP.AddItem "CONSOL CON AJUSTE                        7"
    cboMonedaP.ListIndex = 0
End Sub

Private Sub Limpiar()
Dim m As Integer
fgPresu.MergeCells = flexMergeFree
If Right(Trim(cboTpo.Text), 1) = "2" Then
    'PRESUPUESTO
    Call MSHFlex(fgPresu, 62, "Item-Código-Descripción-Año " & Val(Trim(cboFecha.Text)) - 1 & "-Presupuesto-Cred.1-Cred.2-Total-Var.Monto-Var. %-Enero-Enero-Enero-Enero-Febrero-Febrero-Febrero-Febrero-Marzo-Marzo-Marzo-Marzo-Abril-Abril-Abril-Abril-Mayo-Mayo-Mayo-Mayo-Junio-Junio-Junio-Junio-Julio-Julio-Julio-Julio-Agosto-Agosto-Agosto-Agosto-Setiembre-Setiembre-Setiembre-Setiembre-Octubre-Octubre-Octubre-Octubre-Noviembre-Noviembre-Noviembre-Noviembre-Diciembre-Diciembre-Diciembre-Diciembre-TOTALES-TOTALES-TOTALES-TOTALES", _
        "400-1150-2100-1100-1100-1000-1000-1100-1100-700-1100-1100-1100-0-1100-1100-1100-0-1100-1100-1100-0-1100-1100-1100-0-1100-1100-1100-0-1100-1100-1100-0-1100-1100-1100-0-1100-1100-1100-0-1100-1100-1100-0-1100-1100-1100-0-1100-1100-1100-0-1100-1100-1100-0-0-0-0-800", "R-L-L-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-C")
Else
    Call MSHFlex(fgPresu, 62, "Item-Código-Descripción-Año " & Val(Trim(cboFecha.Text)) - 1 & "-Presupuesto-Cred.1-Cred.2-Total-Var.Monto-Var. %-Enero-Enero-Enero-Enero-Febrero-Febrero-Febrero-Febrero-Marzo-Marzo-Marzo-Marzo-Abril-Abril-Abril-Abril-Mayo-Mayo-Mayo-Mayo-Junio-Junio-Junio-Junio-Julio-Julio-Julio-Julio-Agosto-Agosto-Agosto-Agosto-Setiembre-Setiembre-Setiembre-Setiembre-Octubre-Octubre-Octubre-Octubre-Noviembre-Noviembre-Noviembre-Noviembre-Diciembre-Diciembre-Diciembre-Diciembre-TOTALES-TOTALES-TOTALES-TOTALES", _
        "400-1150-2100-1100-1100-1000-1000-1100-1100-700-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1100-1200-1200-1200-800", "R-L-L-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-R-C")
End If
fgPresu.Rows = 2
fgPresu.TextMatrix(1, 0) = "."
fgPresu.TextMatrix(1, 1) = ".."
For m = 1 To 12
    fgPresu.TextMatrix(1, m * 4 + 6) = "Presup.Mes"
    fgPresu.TextMatrix(1, m * 4 + 7) = "Movim.Mes"
    fgPresu.TextMatrix(1, m * 4 + 8) = "Difer.Mes"
    fgPresu.TextMatrix(1, m * 4 + 9) = "Difer.Año"
Next
m = 13
fgPresu.TextMatrix(1, m * 4 + 6) = "Presup.Total"
fgPresu.TextMatrix(1, m * 4 + 7) = "Presup.Movim."
fgPresu.TextMatrix(1, m * 4 + 8) = "Presup.Difer."
fgPresu.TextMatrix(1, m * 4 + 9) = "Avance %"
'fgPresu.Row = 1
'For m = 1 To fgPresu.Cols - 1
'    fgPresu.Col = m
'    fgPresu.CellBackColor = &HC0C000   ' Fondo Titulo
'Next
fgPresu.MergeRow(0) = True
End Sub

Private Sub CargaPresu(Optional ByVal pbEjecutado As Boolean = False, Optional pnTipo As Integer = 1)
Dim tmpReg As New ADODB.Recordset
Dim tmpSql As String
Dim x As Integer, n As Integer
Dim nMonBal As Currency, nMonMov As Currency
Dim TotMov As Currency
Dim oPP As DPresupuesto
Set oPP = New DPresupuesto
Dim nMonSdo As Currency

x = 1
fgPresu.Redraw = False

If Me.cboPresu.Text = "" Then
    Exit Sub
End If

Set tmpReg = oPP.GetPresupuestoEjec(Right(Me.cboPresu.Text, 4), Me.cboFecha.Text, pbEjecutado, chkProy.value, chkBala.value, pnTipo, CInt(Right(Trim(cboTpo.Text), 1)), Right(Me.cboMonedaP.Text, 1), IIf(Me.chkCierreAño.value = 1, True, False), Me.chkMovContables.value)

If Not (tmpReg.BOF Or tmpReg.EOF) Then
    With tmpReg
        Do While Not .EOF
            x = x + 1
            AdicionaRow fgPresu, x
            fgPresu.Row = fgPresu.Rows - 1
            fgPresu.TextMatrix(x, 0) = x - 1
            fgPresu.TextMatrix(x, 1) = !cCodRub
            fgPresu.TextMatrix(x, 2) = !cDesRub
            If !nMonIni <> 0 Or !nMonto <> 0 Then
                fgPresu.TextMatrix(x, 3) = Format(!nMonIni, "#,##0.00")
                fgPresu.TextMatrix(x, 4) = Format(!nMonto, "#,##0.00")
            Else
                fgPresu.TextMatrix(x, 3) = IIf(!nMonIni = 0, "", Format(!nMonIni, "#,##0.00"))
                fgPresu.TextMatrix(x, 4) = IIf(!nMonto = 0, "", Format(!nMonto, "#,##0.00"))
            End If
            fgPresu.TextMatrix(x, 5) = IIf(!nMonCre1 = 0, "", Format(!nMonCre1, "#,##0.00"))
            fgPresu.TextMatrix(x, 6) = IIf(!nMonCre2 = 0, "", Format(!nMonCre2, "#,##0.00"))
            If !nMonIni <> 0 Or !nMonto <> 0 Then
                fgPresu.TextMatrix(x, 7) = Format(!Total, "#,##0.00")
            Else
                fgPresu.TextMatrix(x, 7) = IIf(!Total = 0, "", Format(!Total, "#,##0.00"))
            End If
            If !Total <> 0 And !nMonIni <> 0 Then
                fgPresu.TextMatrix(x, 8) = Format(!Total - !nMonIni, "#,##0.00")
                fgPresu.TextMatrix(x, 9) = Format(Round((((!Total / !nMonIni) - 1) * 100), 2), "#,##0.00")
            ElseIf !Total <> 0 Or !nMonIni <> 0 Then
                fgPresu.TextMatrix(x, 8) = Format(!Total - !nMonIni, "#,##0.00")
                If !Total = 0 Then
                    fgPresu.TextMatrix(x, 9) = Format(100 * -1, "#,##0.00")
                Else
                    fgPresu.TextMatrix(x, 9) = Format(100, "#,##0.00")
                End If
            Else
                fgPresu.TextMatrix(x, 8) = ""
                fgPresu.TextMatrix(x, 9) = ""
            End If
            
            If pbEjecutado Then
                If Right(Trim(cboTpo.Text), 1) = "2" Then
                    
                    If .Fields(0) = "00020101010403" Then
                        MsgBox "ddd"
                        
                    End If
                    
                    
                    'SOLO BALANCES
                    nMonBal = !nMonIni
                    nMonMov = !nMonIni
                    nMonSdo = IIf(IsNull(!nSaldoIni), 0, !nSaldoIni)
                    If Me.chkMensuales.value = 0 Then
                        nMonBal = nMonBal + !Ene: nMonMov = nMonMov + !MovEne
                        nMonSdo = nMonSdo + !MovEne
                    Else
                        nMonBal = !Ene: nMonMov = !MovEne
                        nMonSdo = !Ene - !MovEne
                    End If
                    
                    fgPresu.TextMatrix(x, 10) = IIf(!Ene = 0, "", Format(nMonBal, "#,##0.00"))
                    fgPresu.TextMatrix(x, 11) = IIf(!MovEne = 0, "", Format(nMonMov, "#,##0.00"))
                    fgPresu.TextMatrix(x, 12) = IIf(!Ene = 0, "", Format(nMonSdo, "#,##0.00"))
                    'fgPresu.TextMatrix(x, 13) = "" 'IIf(!DMEne = 0, "", Format(!DMEne, "#,##0.00"))
                    If Me.chkMovContables.value = 0 Then
                        fgPresu.TextMatrix(x, 13) = "" 'IIf(!DMEne + !DMFeb = 0, "", Format(!DMEne + !DMFeb, "#,##0.00"))
                        Me.fgPresu.ColWidth(13) = 0
                    Else
                        fgPresu.TextMatrix(x, 13) = IIf(!MovCEne = 0, "", Format(!MovCEne, "#,##0.00"))
                        Me.fgPresu.ColWidth(13) = 1095
                        fgPresu.TextMatrix(1, 13) = "Mov.Cont.Mes"
                    End If
                    If Me.chkMensuales.value = 0 Then
                        nMonBal = nMonBal + !Feb: nMonMov = nMonMov + !MovFEb
                        nMonSdo = nMonSdo + !MovFEb
                    Else
                        nMonBal = !Feb: nMonMov = !MovFEb
                        nMonSdo = !Feb - !MovFEb
                    End If
                    fgPresu.TextMatrix(x, 14) = IIf(!Feb = 0, "", Format(nMonBal, "#,##0.00"))
                    fgPresu.TextMatrix(x, 15) = IIf(!MovFEb = 0, "", Format(nMonMov, "#,##0.00"))
                    fgPresu.TextMatrix(x, 16) = IIf(!Feb = 0, "", Format(nMonSdo, "#,##0.00"))
                    If Me.chkMovContables.value = 0 Then
                        fgPresu.TextMatrix(x, 17) = "" 'IIf(!DMEne + !DMFeb = 0, "", Format(!DMEne + !DMFeb, "#,##0.00"))
                        Me.fgPresu.ColWidth(17) = 0
                    Else
                        fgPresu.TextMatrix(x, 17) = IIf(!MovCFEb = 0, "", Format(!MovCFEb, "#,##0.00"))
                        Me.fgPresu.ColWidth(17) = 1095
                        fgPresu.TextMatrix(1, 17) = "Mov.Cont.Mes"
                    End If
                    
                    If Me.chkMensuales.value = 0 Then
                        nMonBal = nMonBal + !Mar: nMonMov = nMonMov + !MovMar
                        nMonSdo = nMonSdo + !MovMar
                    Else
                        nMonBal = !Mar: nMonMov = !MovMar
                        nMonSdo = !Mar - !MovMar
                    End If
                    fgPresu.TextMatrix(x, 18) = IIf(!Mar = 0, "", Format(nMonBal, "#,##0.00"))
                    fgPresu.TextMatrix(x, 19) = IIf(!MovMar = 0, "", Format(nMonMov, "#,##0.00"))
                    fgPresu.TextMatrix(x, 20) = IIf(!Mar = 0, "", Format(nMonSdo, "#,##0.00"))
                    'fgPresu.TextMatrix(x, 21) = "" 'IIf(!DMEne + !DMFeb + !DMMar = 0, "", Format(!DMEne + !DMFeb + !DMMar, "#,##0.00"))
                    If Me.chkMovContables.value = 0 Then
                        fgPresu.TextMatrix(x, 21) = "" 'IIf(!DMEne + !DMFeb = 0, "", Format(!DMEne + !DMFeb, "#,##0.00"))
                        Me.fgPresu.ColWidth(21) = 0
                    Else
                        fgPresu.TextMatrix(x, 21) = IIf(!MovCMar = 0, "", Format(!MovCMar, "#,##0.00"))
                        Me.fgPresu.ColWidth(21) = 1095
                        fgPresu.TextMatrix(1, 21) = "Mov.Cont.Mes"
                    End If
                    If Me.chkMensuales.value = 0 Then
                        nMonBal = nMonBal + !Abr: nMonMov = nMonMov + !MovAbr
                        nMonSdo = nMonSdo + !MovAbr
                    Else
                        nMonBal = !Abr: nMonMov = !MovAbr
                        nMonSdo = !Abr - !MovAbr
                    End If
                    fgPresu.TextMatrix(x, 22) = IIf(!Abr = 0, "", Format(nMonBal, "#,##0.00"))
                    fgPresu.TextMatrix(x, 23) = IIf(!MovAbr = 0, "", Format(nMonMov, "#,##0.00"))
                    fgPresu.TextMatrix(x, 24) = IIf(!Abr = 0, "", Format(nMonSdo, "#,##0.00"))
                    'fgPresu.TextMatrix(x, 25) = "" 'IIf(!DMEne + !DMFeb + !DMMar + !DMAbr = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr, "#,##0.00"))
                    If Me.chkMovContables.value = 0 Then
                        fgPresu.TextMatrix(x, 25) = "" 'IIf(!DMEne + !DMFeb = 0, "", Format(!DMEne + !DMFeb, "#,##0.00"))
                        Me.fgPresu.ColWidth(25) = 0
                    Else
                        fgPresu.TextMatrix(x, 25) = IIf(!MovCAbr = 0, "", Format(!MovCAbr, "#,##0.00"))
                        Me.fgPresu.ColWidth(25) = 1095
                        fgPresu.TextMatrix(1, 25) = "Mov.Cont.Mes"
                    End If
                    If Me.chkMensuales.value = 0 Then
                        nMonBal = nMonBal + !May: nMonMov = nMonMov + !MovMay
                        nMonSdo = nMonSdo + !MovMay
                    Else
                        nMonBal = !May: nMonMov = !MovMay
                        nMonSdo = !May - !MovMay
                    End If
                    fgPresu.TextMatrix(x, 26) = IIf(!May = 0, "", Format(nMonBal, "#,##0.00"))
                    fgPresu.TextMatrix(x, 27) = IIf(!MovMay = 0, "", Format(nMonMov, "#,##0.00"))
                    fgPresu.TextMatrix(x, 28) = IIf(!May = 0, "", Format(nMonSdo, "#,##0.00"))
                    'fgPresu.TextMatrix(x, 29) = "" 'IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay, "#,##0.00"))
                    If Me.chkMovContables.value = 0 Then
                        fgPresu.TextMatrix(x, 29) = "" 'IIf(!DMEne + !DMFeb = 0, "", Format(!DMEne + !DMFeb, "#,##0.00"))
                        Me.fgPresu.ColWidth(29) = 0
                    Else
                        fgPresu.TextMatrix(x, 29) = IIf(!MovCMay = 0, "", Format(!MovCMay, "#,##0.00"))
                        Me.fgPresu.ColWidth(29) = 1095
                        fgPresu.TextMatrix(1, 29) = "Mov.Cont.Mes"
                    End If
                    If Me.chkMensuales.value = 0 Then
                        nMonBal = nMonBal + !Jun: nMonMov = nMonMov + !MovJun
                        nMonSdo = nMonSdo + !MovJun
                    Else
                        nMonBal = !Jun: nMonMov = !MovJun
                        nMonSdo = !Jun - !MovJun
                    End If
                    nMonSdo = nMonSdo + !MovJun
                    fgPresu.TextMatrix(x, 30) = IIf(!Jun = 0, "", Format(nMonBal, "#,##0.00"))
                    fgPresu.TextMatrix(x, 31) = IIf(!MovJun = 0, "", Format(nMonMov, "#,##0.00"))
                    fgPresu.TextMatrix(x, 32) = IIf(!Jun = 0, "", Format(nMonSdo, "#,##0.00"))
                    'fgPresu.TextMatrix(x, 33) = "" 'IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun, "#,##0.00"))
                    If Me.chkMovContables.value = 0 Then
                        fgPresu.TextMatrix(x, 33) = "" 'IIf(!DMEne + !DMFeb = 0, "", Format(!DMEne + !DMFeb, "#,##0.00"))
                        Me.fgPresu.ColWidth(33) = 0
                    Else
                        fgPresu.TextMatrix(x, 33) = IIf(!MovCJun = 0, "", Format(!MovCJun, "#,##0.00"))
                        Me.fgPresu.ColWidth(33) = 1095
                        fgPresu.TextMatrix(1, 33) = "Mov.Cont.Mes"
                    End If
                    If Me.chkMensuales.value = 0 Then
                        nMonBal = nMonBal + !Jul: nMonMov = nMonMov + !MovJul
                        nMonSdo = nMonSdo + !MovJul
                    Else
                        nMonBal = !Jul: nMonMov = !MovJul
                        nMonSdo = !Jul - !MovJul
                    End If
                    fgPresu.TextMatrix(x, 34) = IIf(!Jul = 0, "", Format(nMonBal, "#,##0.00"))
                    fgPresu.TextMatrix(x, 35) = IIf(!MovJul = 0, "", Format(nMonMov, "#,##0.00"))
                    fgPresu.TextMatrix(x, 36) = IIf(!Jul = 0, "", Format(nMonSdo, "#,##0.00"))
                    'fgPresu.TextMatrix(x, 37) = "" 'IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul, "#,##0.00"))
                    If Me.chkMovContables.value = 0 Then
                        fgPresu.TextMatrix(x, 37) = "" 'IIf(!DMEne + !DMFeb = 0, "", Format(!DMEne + !DMFeb, "#,##0.00"))
                        Me.fgPresu.ColWidth(37) = 0
                    Else
                        fgPresu.TextMatrix(x, 37) = IIf(!MovCJul = 0, "", Format(!MovCJul, "#,##0.00"))
                        Me.fgPresu.ColWidth(37) = 1095
                        fgPresu.TextMatrix(1, 37) = "Mov.Cont.Mes"
                    End If
                    If Me.chkMensuales.value = 0 Then
                        nMonBal = nMonBal + !Ago: nMonMov = nMonMov + !MovAgo
                        nMonSdo = nMonSdo + !MovAgo
                    Else
                        nMonBal = !Ago: nMonMov = !MovAgo
                        nMonSdo = !Ago - !MovAgo
                    End If
                    fgPresu.TextMatrix(x, 38) = IIf(!Ago = 0, "", Format(nMonBal, "#,##0.00"))
                    fgPresu.TextMatrix(x, 39) = IIf(!MovAgo = 0, "", Format(nMonMov, "#,##0.00"))
                    fgPresu.TextMatrix(x, 40) = IIf(!Ago = 0, "", Format(nMonSdo, "#,##0.00"))
                    'fgPresu.TextMatrix(x, 41) = "" 'IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo, "#,##0.00"))
                    If Me.chkMovContables.value = 0 Then
                        fgPresu.TextMatrix(x, 41) = "" 'IIf(!DMEne + !DMFeb = 0, "", Format(!DMEne + !DMFeb, "#,##0.00"))
                        Me.fgPresu.ColWidth(41) = 0
                    Else
                        fgPresu.TextMatrix(x, 41) = IIf(!MovCAgo = 0, "", Format(!MovCAgo, "#,##0.00"))
                        Me.fgPresu.ColWidth(41) = 1095
                        fgPresu.TextMatrix(1, 41) = "Mov.Cont.Mes"
                    End If
                    If Me.chkMensuales.value = 0 Then
                        nMonBal = nMonBal + !Seti: nMonMov = nMonMov + !MovSet
                        nMonSdo = nMonSdo + !MovSet
                    Else
                        nMonBal = !Seti: nMonMov = !MovSet
                        nMonSdo = !Seti - !MovSet
                    End If
                    fgPresu.TextMatrix(x, 42) = IIf(!Seti = 0, "", Format(nMonBal, "#,##0.00"))
                    fgPresu.TextMatrix(x, 43) = IIf(!MovSet = 0, "", Format(nMonMov, "#,##0.00"))
                    fgPresu.TextMatrix(x, 44) = IIf(!Seti = 0, "", Format(nMonSdo, "#,##0.00"))
                    'fgPresu.TextMatrix(x, 45) = "" 'IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet, "#,##0.00"))
                    If Me.chkMovContables.value = 0 Then
                        fgPresu.TextMatrix(x, 45) = "" 'IIf(!DMEne + !DMFeb = 0, "", Format(!DMEne + !DMFeb, "#,##0.00"))
                        Me.fgPresu.ColWidth(45) = 0
                    Else
                        fgPresu.TextMatrix(x, 45) = IIf(!MovCSet = 0, "", Format(!MovCSet, "#,##0.00"))
                        Me.fgPresu.ColWidth(45) = 1095
                        fgPresu.TextMatrix(1, 45) = "Mov.Cont.Mes"
                    End If
                    If Me.chkMensuales.value = 0 Then
                        nMonBal = nMonBal + !Oct: nMonMov = nMonMov + !MovOct
                        nMonSdo = nMonSdo + !MovOct
                    Else
                        nMonBal = !Oct: nMonMov = !MovOct
                        nMonSdo = !Oct - !MovOct
                    End If
                    fgPresu.TextMatrix(x, 46) = IIf(!Oct = 0, "", Format(nMonBal, "#,##0.00"))
                    fgPresu.TextMatrix(x, 47) = IIf(!MovOct = 0, "", Format(nMonMov, "#,##0.00"))
                    fgPresu.TextMatrix(x, 48) = IIf(!Oct = 0, "", Format(nMonSdo, "#,##0.00"))
                    'fgPresu.TextMatrix(x, 49) = "" 'IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet + !DMOct = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet + !DMOct, "#,##0.00"))
                    If Me.chkMovContables.value = 0 Then
                        fgPresu.TextMatrix(x, 49) = "" 'IIf(!DMEne + !DMFeb = 0, "", Format(!DMEne + !DMFeb, "#,##0.00"))
                        Me.fgPresu.ColWidth(49) = 0
                    Else
                        fgPresu.TextMatrix(x, 49) = IIf(!MovCOct = 0, "", Format(!MovCOct, "#,##0.00"))
                        Me.fgPresu.ColWidth(49) = 1095
                        fgPresu.TextMatrix(1, 49) = "Mov.Cont.Mes"
                    End If
                    If Me.chkMensuales.value = 0 Then
                        nMonBal = nMonBal + !Nov: nMonMov = nMonMov + !MovNov
                        nMonSdo = nMonSdo + !MovNov
                    Else
                        nMonBal = !Nov: nMonMov = !MovNov
                        nMonSdo = !Nov - !MovNov
                    End If
                    fgPresu.TextMatrix(x, 50) = IIf(!Nov = 0, "", Format(nMonBal, "#,##0.00"))
                    fgPresu.TextMatrix(x, 51) = IIf(!MovNov = 0, "", Format(nMonMov, "#,##0.00"))
                    fgPresu.TextMatrix(x, 52) = IIf(!Nov = 0, "", Format(nMonSdo, "#,##0.00"))
                    'fgPresu.TextMatrix(x, 53) = "" 'IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet + !DMOct + !DMNov = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet + !DMOct + !DMNov, "#,##0.00"))
                    If Me.chkMovContables.value = 0 Then
                        fgPresu.TextMatrix(x, 53) = "" 'IIf(!DMEne + !DMFeb = 0, "", Format(!DMEne + !DMFeb, "#,##0.00"))
                        Me.fgPresu.ColWidth(53) = 0
                    Else
                        fgPresu.TextMatrix(x, 53) = IIf(!MovCNov = 0, "", Format(!MovCNov, "#,##0.00"))
                        Me.fgPresu.ColWidth(53) = 1095
                        fgPresu.TextMatrix(1, 53) = "Mov.Cont.Mes"
                    End If
                    If Me.chkMensuales.value = 0 Then
                        nMonBal = nMonBal + !Dic: nMonMov = nMonMov + !MovDic
                       nMonSdo = nMonSdo + !MovDic
                    Else
                        nMonBal = !Dic: nMonMov = !MovDic
                        nMonSdo = !Dic - !MovDic
                    End If
                    fgPresu.TextMatrix(x, 54) = IIf(!Dic = 0, "", Format(nMonBal, "#,##0.00"))
                    fgPresu.TextMatrix(x, 55) = IIf(!MovDic = 0, "", Format(nMonMov, "#,##0.00"))
                    fgPresu.TextMatrix(x, 56) = IIf(!Dic = 0, "", Format(nMonSdo, "#,##0.00"))
                    'fgPresu.TextMatrix(x, 57) = "" 'IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet + !DMOct + !DMNov + !DMDic = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet + !DMOct + !DMNov + !DMDic, "#,##0.00"))
                    If Me.chkMovContables.value = 0 Then
                        fgPresu.TextMatrix(x, 57) = "" 'IIf(!DMEne + !DMFeb = 0, "", Format(!DMEne + !DMFeb, "#,##0.00"))
                        Me.fgPresu.ColWidth(57) = 0
                    Else
                        fgPresu.TextMatrix(x, 57) = IIf(!MovCDic = 0, "", Format(!MovCDic, "#,##0.00"))
                        Me.fgPresu.ColWidth(57) = 1095
                        fgPresu.TextMatrix(1, 57) = "Mov.Cont.Mes"
                    End If
    
                    If nMonBal > 0 And nMonMov > 0 And Me.chkMensuales.value = 0 Then
                        fgPresu.TextMatrix(x, 61) = Round(((nMonMov / nMonBal) * 100), 2)
                    Else
                        fgPresu.TextMatrix(x, 61) = ""
                    End If
                Else
                    fgPresu.TextMatrix(x, 10) = IIf(!Ene = 0, "", Format(!Ene, "#,##0.00"))
                    fgPresu.TextMatrix(x, 11) = IIf(!MovEne = 0, "", Format(!MovEne, "#,##0.00"))
                    fgPresu.TextMatrix(x, 12) = IIf(!DMEne = 0, "", Format(!DMEne, "#,##0.00"))
                    fgPresu.TextMatrix(x, 13) = IIf(!DMEne = 0, "", Format(!DMEne, "#,##0.00"))
                    fgPresu.TextMatrix(x, 14) = IIf(!Feb = 0, "", Format(!Feb, "#,##0.00"))
                    fgPresu.TextMatrix(x, 15) = IIf(!MovFEb = 0, "", Format(!MovFEb, "#,##0.00"))
                    fgPresu.TextMatrix(x, 16) = IIf(!DMFeb = 0, "", Format(!DMFeb, "#,##0.00"))
                    fgPresu.TextMatrix(x, 17) = IIf(!DMEne + !DMFeb = 0, "", Format(!DMEne + !DMFeb, "#,##0.00"))
                    fgPresu.TextMatrix(x, 18) = IIf(!Mar = 0, "", Format(!Mar, "#,##0.00"))
                    fgPresu.TextMatrix(x, 19) = IIf(!MovMar = 0, "", Format(!MovMar, "#,##0.00"))
                    fgPresu.TextMatrix(x, 20) = IIf(!DMMar = 0, "", Format(!DMMar, "#,##0.00"))
                    fgPresu.TextMatrix(x, 21) = IIf(!DMEne + !DMFeb + !DMMar = 0, "", Format(!DMEne + !DMFeb + !DMMar, "#,##0.00"))
                    fgPresu.TextMatrix(x, 22) = IIf(!Abr = 0, "", Format(!Abr, "#,##0.00"))
                    fgPresu.TextMatrix(x, 23) = IIf(!MovAbr = 0, "", Format(!MovAbr, "#,##0.00"))
                    fgPresu.TextMatrix(x, 24) = IIf(!DMAbr = 0, "", Format(!DMAbr, "#,##0.00"))
                    fgPresu.TextMatrix(x, 25) = IIf(!DMEne + !DMFeb + !DMMar + !DMAbr = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr, "#,##0.00"))
                    fgPresu.TextMatrix(x, 26) = IIf(!May = 0, "", Format(!May, "#,##0.00"))
                    fgPresu.TextMatrix(x, 27) = IIf(!MovMay = 0, "", Format(!MovMay, "#,##0.00"))
                    fgPresu.TextMatrix(x, 28) = IIf(!DMMay = 0, "", Format(!DMMay, "#,##0.00"))
                    fgPresu.TextMatrix(x, 29) = IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay, "#,##0.00"))
                    fgPresu.TextMatrix(x, 30) = IIf(!Jun = 0, "", Format(!Jun, "#,##0.00"))
                    fgPresu.TextMatrix(x, 31) = IIf(!MovJun = 0, "", Format(!MovJun, "#,##0.00"))
                    fgPresu.TextMatrix(x, 32) = IIf(!DMJun = 0, "", Format(!DMJun, "#,##0.00"))
                    fgPresu.TextMatrix(x, 33) = IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun, "#,##0.00"))
                    fgPresu.TextMatrix(x, 34) = IIf(!Jul = 0, "", Format(!Jul, "#,##0.00"))
                    fgPresu.TextMatrix(x, 35) = IIf(!MovJul = 0, "", Format(!MovJul, "#,##0.00"))
                    fgPresu.TextMatrix(x, 36) = IIf(!DMJul = 0, "", Format(!DMJul, "#,##0.00"))
                    fgPresu.TextMatrix(x, 37) = IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul, "#,##0.00"))
                    fgPresu.TextMatrix(x, 38) = IIf(!Ago = 0, "", Format(!Ago, "#,##0.00"))
                    fgPresu.TextMatrix(x, 39) = IIf(!MovAgo = 0, "", Format(!MovAgo, "#,##0.00"))
                    fgPresu.TextMatrix(x, 40) = IIf(!DMAgo = 0, "", Format(!DMAgo, "#,##0.00"))
                    fgPresu.TextMatrix(x, 41) = IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo, "#,##0.00"))
                    fgPresu.TextMatrix(x, 42) = IIf(!Seti = 0, "", Format(!Seti, "#,##0.00"))
                    fgPresu.TextMatrix(x, 43) = IIf(!MovSet = 0, "", Format(!MovSet, "#,##0.00"))
                    fgPresu.TextMatrix(x, 44) = IIf(!DMSet = 0, "", Format(!DMSet, "#,##0.00"))
                    fgPresu.TextMatrix(x, 45) = IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet, "#,##0.00"))
                    fgPresu.TextMatrix(x, 46) = IIf(!Oct = 0, "", Format(!Oct, "#,##0.00"))
                    fgPresu.TextMatrix(x, 47) = IIf(!MovOct = 0, "", Format(!MovOct, "#,##0.00"))
                    fgPresu.TextMatrix(x, 48) = IIf(!DMOct = 0, "", Format(!DMOct, "#,##0.00"))
                    fgPresu.TextMatrix(x, 49) = IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet + !DMOct = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet + !DMOct, "#,##0.00"))
                    fgPresu.TextMatrix(x, 50) = IIf(!Nov = 0, "", Format(!Nov, "#,##0.00"))
                    fgPresu.TextMatrix(x, 51) = IIf(!MovNov = 0, "", Format(!MovNov, "#,##0.00"))
                    fgPresu.TextMatrix(x, 52) = IIf(!DMNov = 0, "", Format(!DMNov, "#,##0.00"))
                    fgPresu.TextMatrix(x, 53) = IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet + !DMOct + !DMNov = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet + !DMOct + !DMNov, "#,##0.00"))
                    fgPresu.TextMatrix(x, 54) = IIf(!Dic = 0, "", Format(!Dic, "#,##0.00"))
                    fgPresu.TextMatrix(x, 55) = IIf(!MovDic = 0, "", Format(!MovDic, "#,##0.00"))
                    fgPresu.TextMatrix(x, 56) = IIf(!DMDic = 0, "", Format(!DMDic, "#,##0.00"))
                    fgPresu.TextMatrix(x, 57) = IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet + !DMOct + !DMNov + !DMDic = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet + !DMOct + !DMNov + !DMDic, "#,##0.00"))
                
                    'Totales
                    TotMov = !MovEne + !MovFEb + !MovMar + !MovAbr + !MovMay + !MovJun + !MovJul + !MovAgo + !MovSet + !MovOct + !MovNov + !MovDic
                    fgPresu.TextMatrix(x, 58) = IIf(!Total = 0, "", Format(!Total, "#,##0.00"))
                    fgPresu.TextMatrix(x, 59) = IIf(TotMov = 0, "", Format(TotMov, "#,##0.00"))
                    fgPresu.TextMatrix(x, 60) = IIf(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet + !DMOct + !DMNov + !DMDic = 0, "", Format(!DMEne + !DMFeb + !DMMar + !DMAbr + !DMMay + !DMJun + !DMJul + !DMAgo + !DMSet + !DMOct + !DMNov + !DMDic, "#,##0.00"))
                    If !Total > 0 And TotMov > 0 Then
                        fgPresu.TextMatrix(x, 61) = Round((((TotMov / !Total)) * 100), 2)
                    Else
                        fgPresu.TextMatrix(x, 61) = ""
                    End If
                End If
            End If
            .MoveNext
        Loop
    End With
End If
tmpReg.Close
Set tmpReg = Nothing
fgPresu.Redraw = True
End Sub

Private Sub Cabecera(ByVal cTipo As String, ByVal nPage As Integer)
    If nPage > 1 Then vRTFImp = vRTFImp & oImpresora.gPrnSaltoPagina
    vRTFImp = vRTFImp & "  CMAC - TRUJILLO" & Space(90) & Format(gdFecSis & " " & Time, gsFormatoFechaHoraView) & oImpresora.gPrnSaltoLinea
    vRTFImp = vRTFImp & ImpreFormat(UCase(gsNomAge), 25) & Space(85) & " Página :" & ImpreFormat(nPage, 5, 0) & oImpresora.gPrnSaltoLinea
    If cTipo = "1" Then
        vRTFImp = vRTFImp & Space(30) & "LISTADO DEL AÑO " & Left(Trim(cboFecha), 4) & " DE : " & UCase(Left(cboPresu, 25)) & " EN " & Left(Trim(cboMoneda), 12) & oImpresora.gPrnSaltoLinea
    End If
    vRTFImp = vRTFImp & Space(2) & String(124, "-") & oImpresora.gPrnSaltoLinea
    vRTFImp = vRTFImp & Space(2) & "    CODIGO           DESCRIPCION          PERIODO         PRESUP. MES       MOVIM. MES       DIFER. MES       DIFER. AÑO" & oImpresora.gPrnSaltoLinea
    vRTFImp = vRTFImp & Space(2) & String(124, "-") & oImpresora.gPrnSaltoLinea
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CierraConexion
End Sub


