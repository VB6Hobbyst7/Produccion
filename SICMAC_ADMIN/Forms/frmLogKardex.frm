VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLogKardex 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kardex de Productos"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11565
   Icon            =   "frmLogKardex.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTarjeta 
      Caption         =   "Tarjeta"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6720
      TabIndex        =   24
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdExportarExcel 
      Caption         =   "&Exportar"
      Height          =   360
      Left            =   4935
      TabIndex        =   23
      Top             =   5160
      Width           =   1140
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   360
      Left            =   9135
      TabIndex        =   6
      Top             =   5160
      Width           =   1140
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   360
      Left            =   7935
      TabIndex        =   5
      Top             =   5160
      Width           =   1140
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   10335
      TabIndex        =   7
      Top             =   5160
      Width           =   1140
   End
   Begin Sicmact.FlexEdit Flex 
      Height          =   3915
      Left            =   45
      TabIndex        =   8
      Top             =   1080
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   6906
      Cols0           =   13
      EncabezadosNombres=   "#-Fecha-Documento-# Ingreso-# Salida-# Saldo-Proveedor-Age-Area-Pre.Prom.-Val Ingreso-Val Salida-Saldo"
      EncabezadosAnchos=   "400-1200-1200-1500-1500-1500-3000-800-1500-1200-1200-1200-1200"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-L-R-R-R-R-L-R-L-R-R-R-R"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
      TextArray0      =   "#"
      Appearance      =   0
      ColWidth0       =   405
      RowHeight0      =   300
   End
   Begin VB.Frame fraProducto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Producto"
      ForeColor       =   &H00800000&
      Height          =   990
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   11400
      Begin VB.ComboBox cboTpoAlm 
         Height          =   315
         Left            =   8235
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   255
         Width           =   3090
      End
      Begin Sicmact.TxtBuscar txtProducto 
         Height          =   300
         Left            =   105
         TabIndex        =   0
         Top             =   270
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   529
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
      End
      Begin MSMask.MaskEdBox mskFecIni 
         Height          =   300
         Left            =   7695
         TabIndex        =   3
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFecFin 
         Height          =   300
         Left            =   9960
         TabIndex        =   4
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Sicmact.TxtBuscar txtAlmacen 
         Height          =   300
         Left            =   855
         TabIndex        =   2
         Top             =   600
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
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
         Left            =   6885
         TabIndex        =   22
         Top             =   322
         Width           =   1500
      End
      Begin VB.Label lblAlmacen 
         Caption         =   "Almacen"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   135
         TabIndex        =   13
         Top             =   645
         Width           =   705
      End
      Begin VB.Label lblAlmacenG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2025
         TabIndex        =   12
         Top             =   615
         Width           =   4650
      End
      Begin VB.Label lblFecFin 
         Caption         =   "Fecha Fin."
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   9075
         TabIndex        =   11
         Top             =   653
         Width           =   915
      End
      Begin VB.Label lblFecIni 
         Caption         =   "Fecha Ini."
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   6900
         TabIndex        =   10
         Top             =   653
         Width           =   915
      End
      Begin VB.Label lblProducto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2025
         TabIndex        =   9
         Top             =   270
         Width           =   4635
      End
   End
   Begin VB.Label lblSaldos 
      Caption         =   "Saldos"
      Height          =   195
      Left            =   3240
      TabIndex        =   20
      Top             =   5040
      Width           =   720
   End
   Begin VB.Label Label5 
      Caption         =   "="
      Height          =   210
      Left            =   2970
      TabIndex        =   19
      Top             =   5295
      Width           =   225
   End
   Begin VB.Label lblSaldosG 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
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
      Height          =   270
      Left            =   3240
      TabIndex        =   18
      Top             =   5250
      Width           =   1365
   End
   Begin VB.Label lblSalG 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
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
      Height          =   270
      Left            =   1485
      TabIndex        =   16
      Top             =   5250
      Width           =   1365
   End
   Begin VB.Label lblIngresos 
      Caption         =   "Salidas :"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   1
      Left            =   1500
      TabIndex        =   17
      Top             =   5025
      Width           =   810
   End
   Begin VB.Label lblIngG 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   270
      Left            =   45
      TabIndex        =   15
      Top             =   5250
      Width           =   1365
   End
   Begin VB.Label lblIngresos 
      Caption         =   "Ingresos :"
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   0
      Left            =   45
      TabIndex        =   14
      Top             =   5010
      Width           =   810
   End
End
Attribute VB_Name = "frmLogKardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
'*******************************


Private Sub cmdExportarExcel_Click()
Dim rs As ADODB.Recordset

Dim fs              As Scripting.FileSystemObject
Dim xlAplicacion    As Excel.Application
Dim xlLibro         As Excel.Workbook
Dim xlHoja1         As Excel.Worksheet
Dim lbExisteHoja    As Boolean
Dim liLineas        As Integer
Dim i               As Integer
Dim glsArchivo      As String
Dim lsNomHoja       As String

Dim lsFecha As String
Dim lsDocumento As String
Dim lsIngreso As String
Dim lsSalida As String
Dim lsSaldo As String
Dim lsArea As String
Dim lsAge As String
Dim lsValIngreso As String
Dim lsValSalida As String
Dim lsSaldoMont As String
Dim lsIngTotal As String
Dim lsSalTotal As String
Dim lsProveedor As String

Dim lnIngTot As Currency
Dim lnSalTot As Currency
Dim lnIngTotal As Currency
Dim lnSalTotal As Currency
Dim oCon As DConecta
Set oCon = New DConecta

    If Me.flex.TextMatrix(1, 1) = "" Then
        MsgBox "Debe procesar previamente el Kardex antes de Exportarlo a EXCEL.", vbInformation, "Aviso"
        Me.cmdProcesar.SetFocus
        Exit Sub
    End If
        
    glsArchivo = "Reporte_KARDEX_Bienes" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
    Set fs = New Scripting.FileSystemObject

    Set xlAplicacion = New Excel.Application
    If fs.FileExists(App.path & "\SPOOLER\" & glsArchivo) Then
        Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & glsArchivo)
    Else
        Set xlLibro = xlAplicacion.Workbooks.Add
    End If
    Set xlHoja1 = xlLibro.Worksheets.Add

    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 60
    xlHoja1.PageSetup.Orientation = xlLandscape

            lbExisteHoja = False
            lsNomHoja = "KARDEX Bienes"
            For Each xlHoja1 In xlLibro.Worksheets
                If xlHoja1.Name = lsNomHoja Then
                    xlHoja1.Activate
                    lbExisteHoja = True
                    Exit For
                End If
            Next
            If lbExisteHoja = False Then
                Set xlHoja1 = xlLibro.Worksheets.Add
                xlHoja1.Name = lsNomHoja
            End If

            xlAplicacion.Range("A1:A1").ColumnWidth = 10
            xlAplicacion.Range("B1:B1").ColumnWidth = 15
            xlAplicacion.Range("c1:c1").ColumnWidth = 10
            xlAplicacion.Range("D1:D1").ColumnWidth = 10
            xlAplicacion.Range("E1:E1").ColumnWidth = 12
            xlAplicacion.Range("F1:F1").ColumnWidth = 30
            xlAplicacion.Range("G1:G1").ColumnWidth = 30
            xlAplicacion.Range("H1:H1").ColumnWidth = 30
            xlAplicacion.Range("I1:I1").ColumnWidth = 6
            xlAplicacion.Range("J1:J1").ColumnWidth = 12
            xlAplicacion.Range("K1:K1").ColumnWidth = 12
            xlAplicacion.Range("L1:L1").ColumnWidth = 12
            xlAplicacion.Range("M1:M1").ColumnWidth = 12
            xlAplicacion.Range("L1:J1").ColumnWidth = 12
                    
            xlAplicacion.Range("A1:Z100").Font.Size = 9
       
            xlHoja1.Cells(1, 1) = "CAJA MUNICIPAL DE MAYNAS"
            xlHoja1.Cells(2, 1) = "Activos Fijos"
            xlHoja1.Cells(3, 7) = "Fecha :" & Format(gdFecSis, "dd/mm/yyyy")
            xlHoja1.Cells(5, 4) = "KARDEX"
            
                      
            xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 3)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(5, 4), xlHoja1.Cells(5, 8)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 3)).Merge True
            xlHoja1.Range(xlHoja1.Cells(5, 4), xlHoja1.Cells(5, 8)).Merge True
           
            xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 12)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 12)).Borders(xlEdgeBottom).Weight = xlMedium
            xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 12)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
            
            xlHoja1.Cells(6, 1) = "Ubicacion : " & Trim(Me.lblAlmacenG.Caption)
            xlHoja1.Cells(7, 1) = "Descripcion : " & Trim(Me.lblProducto.Caption)
                                  
            liLineas = 8
            
            xlHoja1.Cells(liLineas, 1) = "Fecha"
            xlHoja1.Cells(liLineas, 2) = "Documento"
            xlHoja1.Cells(liLineas, 3) = "Ingreso"
            xlHoja1.Cells(liLineas, 4) = "Salida"
            xlHoja1.Cells(liLineas, 5) = "Saldo"
            xlHoja1.Cells(liLineas, 6) = "Proveedor" 'EJVG20111115
            xlHoja1.Cells(liLineas, 7) = "Area"
            xlHoja1.Cells(liLineas, 8) = "Agencia" 'EJVG20111115
            xlHoja1.Cells(liLineas, 9) = "AGE"
            xlHoja1.Cells(liLineas, 10) = "Val Ingreso"
            xlHoja1.Cells(liLineas, 11) = "Val Salida"
            xlHoja1.Cells(liLineas, 12) = "Saldo"
           
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 12)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 12)).VerticalAlignment = xlCenter

            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas + 2, 1)).Merge True
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 12)).EntireRow.AutoFit
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 12)).WrapText = True
            
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 12)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 12)).Borders.LineStyle = 1
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 12)).Interior.Color = RGB(159, 206, 238)
            
                   
            liLineas = liLineas + 1
                 
      Dim lnI As Integer
      
      For lnI = 1 To Me.flex.Rows - 1
         lsFecha = Me.flex.TextMatrix(lnI, 1)
         lsDocumento = Me.flex.TextMatrix(lnI, 2)
         lsIngreso = Me.flex.TextMatrix(lnI, 3)
         lsSalida = Me.flex.TextMatrix(lnI, 4)
         lsSaldo = Me.flex.TextMatrix(lnI, 5)
         lsProveedor = Me.flex.TextMatrix(lnI, 6)
         lsArea = Me.flex.TextMatrix(lnI, 8)
         lsAge = Me.flex.TextMatrix(lnI, 7)
         lsValIngreso = Me.flex.TextMatrix(lnI, 10)
         lsValSalida = Me.flex.TextMatrix(lnI, 11)
         lsSaldoMont = Me.flex.TextMatrix(lnI, 12)

         lnIngTot = lnIngTot + IIf(Trim(lsIngreso) = "", 0, lsIngreso)
         lnSalTot = lnSalTot + IIf(Trim(lsSalida) = "", 0, lsSalida)
         
         lnIngTotal = lnIngTotal + IIf(Trim(lsValIngreso) = "", 0, lsValIngreso)
         lnSalTotal = lnSalTotal + IIf(Trim(lsValSalida) = "", 0, lsValSalida)
                
            xlHoja1.Cells(liLineas, 1) = lsFecha
            xlHoja1.Cells(liLineas, 2) = lsDocumento
            xlHoja1.Cells(liLineas, 3) = lsIngreso
            xlHoja1.Cells(liLineas, 4) = lsSalida
            xlHoja1.Cells(liLineas, 5) = lsSaldo
            xlHoja1.Cells(liLineas, 6) = lsProveedor
            xlHoja1.Cells(liLineas, 7) = lsArea
            xlHoja1.Cells(liLineas, 8) = IIf(lsAge <> "", GetAgencias(lsAge), "")
            xlHoja1.Cells(liLineas, 9) = lsAge
            xlHoja1.Cells(liLineas, 10) = lsValIngreso
            xlHoja1.Cells(liLineas, 11) = lsValSalida
            xlHoja1.Cells(liLineas, 12) = lsSaldoMont
              
        
        liLineas = liLineas + 1
    Next lnI
    
    xlHoja1.Cells(liLineas, 3) = Val(lnIngTot)
    xlHoja1.Cells(liLineas, 4) = Val(lnSalTot)
    xlHoja1.Cells(liLineas, 5) = Val(lnIngTot) - Val(lnSalTot)
    xlHoja1.Cells(liLineas, 10) = lnIngTotal
    xlHoja1.Cells(liLineas, 11) = lnSalTotal
    xlHoja1.Cells(liLineas, 12) = lnIngTotal + lnSalTotal
    
    xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 12)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 12)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 12)).Borders.LineStyle = 1
    xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 12)).Interior.Color = RGB(159, 206, 238)
            
    
    
   ' ExcelCuadro xlHoja1, 1, 4, 10, liLineas - 1
             
            
'            xlHoja1.Range(xlHoja1.Cells(liLineas, 3), xlHoja1.Cells(liLineas, 5)).Style = "Comma"
'            xlHoja1.Range(xlHoja1.Cells(liLineas, 8), xlHoja1.Cells(liLineas, 8)).Style = "Comma"
'
'            xlHoja1.Range(xlHoja1.Cells(liLineas, 3), xlHoja1.Cells(liLineas, 5)).HorizontalAlignment = xlCenter
'            xlHoja1.Range(xlHoja1.Cells(liLineas, 7), xlHoja1.Cells(liLineas, 7)).HorizontalAlignment = xlCenter
'            xlHoja1.Range(xlHoja1.Cells(liLineas, 10), xlHoja1.Cells(liLineas, 11)).HorizontalAlignment = xlCenter
'
'            xlHoja1.Range(xlHoja1.Cells(liLineas, 6), xlHoja1.Cells(liLineas, 6)).HorizontalAlignment = xlRight
'            xlHoja1.Range(xlHoja1.Cells(liLineas, 10), xlHoja1.Cells(liLineas, 10)).HorizontalAlignment = xlRight
          
                   

        'ExcelCuadro xlHoja1, 1, 4, 12, liLineas - 1
        
        xlHoja1.SaveAs App.path & "\SPOOLER\" & glsArchivo
        ExcelEnd App.path & "\Spooler\" & glsArchivo, xlAplicacion, xlLibro, xlHoja1
    
        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing
        MsgBox "Se ha generado el Archivo en " & App.path & "\SPOOLER\" & glsArchivo
        Call CargaArchivo(glsArchivo, App.path & "\SPOOLER\")
        
        'ARLO 20160126 ***
        gsopecod = LogPistaKardexProducto
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsopecod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Se Genero en Excel el Reporte de Kardex del " & mskFecIni & " al " & mskFecFin
        Set objPista = Nothing
        '**************
 
End Sub

Private Sub CmdImprimir_Click()
    If Me.flex.TextMatrix(1, 1) = "" Then
        MsgBox "Debe procesar previamente el Kardex antes de imprimirlo.", vbInformation, "Aviso"
        Me.cmdProcesar.SetFocus
        Exit Sub
    End If
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim lnI As Long
    Dim oPrevio As clsPrevio
    Set oPrevio = New clsPrevio
    
    Dim lsFecha As String * 10
    Dim lsDocumento As String * 17
    Dim lsIngreso As String * 11
    Dim lsSalida As String * 11
    Dim lsSaldo As String * 11
    Dim lsArea As String * 18
    Dim lsAge As String * 10
    Dim lsValIngreso As String * 14
    Dim lsValSalida As String * 14
    Dim lsSaldoMont As String * 14
    Dim lsIngTotal As String * 14
    Dim lsSalTotal As String * 14
    
    Dim lnIngTotal As Currency
    Dim lnSalTotal As Currency
    
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    lsCadena = ""
    
    lsCadena = lsCadena & oImpresora.gPrnCondensadaON
    
    'lsCadena = lsCadena & CabeceraPagina("KARDEX DE " & Trim(Me.lblProducto.Caption) & " - " & Trim(Me.lblAlmacenG.Caption) & " - " & Me.mskFecIni.Text & " - " & Me.mskFecFin.Text, lnPagina, lnItem, gsNomAge, gsEmpresa, gdFecSis, "")
    'lsCadena = lsCadena & Encabezado("Fecha;8;Documento;18;Ingreso;11;Salida;11;Saldo;11;Area;10; ;2;AGE;17;Val Ingreso;17;Val Salida ;17;Saldo;10;;2;", lnItem)
    lsCadena = lsCadena & CabeceraPagina1("KARDEX DE " & Trim(Me.lblProducto.Caption) & " - " & Trim(Me.lblAlmacenG.Caption) & " - " & Me.mskFecIni.Text & " - " & Me.mskFecFin.Text, lnPagina, lnItem, gsNomAge, gsEmpresa, gdFecSis, "")
    lsCadena = lsCadena & Encabezado1("Fecha;8;Documento;18;Ingreso;11;Salida;11;Saldo;11;Area;10; ;2;AGE;17;Val Ingreso;17;Val Salida ;17;Saldo;10;;2;", lnItem)
    
    For lnI = 1 To Me.flex.Rows - 1
        RSet lsFecha = Me.flex.TextMatrix(lnI, 1)
        RSet lsDocumento = Me.flex.TextMatrix(lnI, 2)
        RSet lsIngreso = Me.flex.TextMatrix(lnI, 3)
        RSet lsSalida = Me.flex.TextMatrix(lnI, 4)
        RSet lsSaldo = Me.flex.TextMatrix(lnI, 5)
        lsArea = Me.flex.TextMatrix(lnI, 8)
        RSet lsAge = Me.flex.TextMatrix(lnI, 7)
        RSet lsValIngreso = Me.flex.TextMatrix(lnI, 10)
        RSet lsValSalida = Me.flex.TextMatrix(lnI, 11)
        RSet lsSaldoMont = Me.flex.TextMatrix(lnI, 12)

        lnIngTotal = lnIngTotal + IIf(Trim(lsValIngreso) = "", 0, lsValIngreso)
        lnSalTotal = lnSalTotal + IIf(Trim(lsValSalida) = "", 0, lsValSalida)
        lsCadena = lsCadena & Space(5) & lsFecha & lsDocumento & lsIngreso & lsSalida & lsSaldo & "  " & lsArea & lsAge & lsValIngreso & lsValSalida & lsSaldoMont & oImpresora.gPrnSaltoLinea
        
        lnItem = lnItem + 1
        
        If lnItem > 54 Then
            lnItem = 0
            lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
            'lsCadena = lsCadena & CabeceraPagina("KARDEX DE " & Trim(Me.lblProducto.Caption) & " - " & Trim(Me.lblAlmacenG.Caption) & " - " & Me.mskFecIni.Text & " - " & Me.mskFecFin.Text, lnPagina, lnItem, gsNomAge, gsEmpresa, gdFecSis, "")
            'lsCadena = lsCadena & Encabezado("Fecha;8;Documento;18;Ingreso;11;Salida;11;Saldo;11;Area;10; ;2;AGE;17;Val Ingreso;17;Val Salida ;17;Saldo;10;;2;", lnItem)
            lsCadena = lsCadena & CabeceraPagina1("KARDEX DE " & Trim(Me.lblProducto.Caption) & " - " & Trim(Me.lblAlmacenG.Caption) & " - " & Me.mskFecIni.Text & " - " & Me.mskFecFin.Text, lnPagina, lnItem, gsNomAge, gsEmpresa, gdFecSis, "")
            lsCadena = lsCadena & Encabezado1("Fecha;8;Documento;18;Ingreso;11;Salida;11;Saldo;11;Area;10; ;2;AGE;17;Val Ingreso;17;Val Salida ;17;Saldo;10;;2;", lnItem)
        End If
    Next lnI
    lsIngTotal = lnIngTotal
    lsSalTotal = lnSalTotal
    'lsCadena = lsCadena & Encabezado(" ;10; ;13;" & Me.lblIngG.Caption & ";13;" & Me.lblSalG.Caption & ";13; ;42; " & lnIngTotal & " ;14;" & lnSalTotal & " ;14;" & lnIngTotal + lnSalTotal & ";13;", lnItem)
    lsCadena = lsCadena & Encabezado1(" ;10; ;13;" & Me.lblIngG.Caption & ";13;" & Me.lblSalG.Caption & ";13; ;42; " & lnIngTotal & " ;14;" & lnSalTotal & " ;14;" & lnIngTotal + lnSalTotal & ";13;", lnItem)
    
    
    oPrevio.Show lsCadena, "KADEX DE " & Trim(Me.lblProducto.Caption) & " - " & Trim(Me.lblAlmacenG.Caption) & " - " & Me.mskFecIni.Text & " - " & Me.mskFecFin.Text, True, 66
    Set oPrevio = Nothing
        
        'ARLO 20160126 ***
        gsopecod = LogPistaKardexProducto
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsopecod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Se Imprimio el Reporte de Kardex del " & mskFecIni & " al " & mskFecFin
        Set objPista = Nothing
        '**************
End Sub


Public Function EncabezadoOrdenCS(psCadena As String, pnItem As Long, Optional pbLineaSimple As Boolean = True) As String
    Dim lsCadena As String
    Dim lsCampo As String
    Dim lnLonCampo As Long
    Dim lnTotalLinea As Long
    Dim lnPos As Long
    Dim lsResultado As String
    Dim i As Long
    Dim lsLineas As String
    
    lsResultado = ""
    lnTotalLinea = 0
        
    lsCadena = psCadena
    pnItem = pnItem + 3
    
    While lsCadena <> ""
        lnPos = InStr(1, lsCadena, ";", vbTextCompare)
        lsCampo = Left(lsCadena, lnPos - 1)
        lsCadena = Mid(lsCadena, lnPos + 1)
        lnPos = InStr(1, lsCadena, ";", vbTextCompare)
        
        lnLonCampo = CCur(Left(lsCadena, lnPos - 1))
        lsCadena = Mid(lsCadena, lnPos + 1)
        lnTotalLinea = lnTotalLinea + lnLonCampo
        lsResultado = lsResultado & Space(lnLonCampo - Len(lsCampo)) & lsCampo
    Wend
        
    lsResultado = lsResultado & oImpresora.gPrnSaltoLinea
    If pbLineaSimple Then
        lsLineas = Space(2) & String(lnTotalLinea + 1, "=") & oImpresora.gPrnSaltoLinea
    Else
        lsLineas = Space(2) & String(lnTotalLinea + 1, "-") & oImpresora.gPrnSaltoLinea
    End If
    
    lsResultado = Space(2) & lsLineas + Space(5) & lsResultado + Space(2) & lsLineas
    
    EncabezadoOrdenCS = lsResultado
End Function

Public Function CabeceraPaginaOrdenCS(ByVal psTitulo As String, pnPagina As Long, pnItem As Long, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date, Optional psMoneda As String = "1") As String
    Dim lsC1 As String
    Dim lsC2 As String
    Dim lsC3 As String
    Dim lsCadena As String
    
    If pnItem >= 66 Then
        pnItem = 0
    End If
    
    pnPagina = pnPagina + 1
    pnItem = 5
    lsCadena = ""

    lsC1 = Format(pgdFecSis, gsFormatoFechaView)
    lsC2 = Format(Time, "hh:mm:ss AMPM")
    lsC3 = "PAGINA Nro. " & Format(pnPagina, "000")
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & pgsEmpresa & Space(39 - Len(lsC3) + 10 - Len(Trim(pgsEmpresa))) & lsC3 & Space(70 - Len(lsC1)) & lsC1 & oImpresora.gPrnSaltoLinea
  
    If psMoneda = "" Then
        lsCadena = lsCadena & Space(5) & pgsNomAge & "      " & Space(114 - Len(lsC2) - Len(lsC2) + 10 - Len(pgsNomAge)) & lsC2 & oImpresora.gPrnSaltoLinea
    ElseIf psMoneda = "1" Then
        '''lsCadena = lsCadena & Space(5) & Trim(pgsNomAge) & "- Soles" + Space(109 - Len("- Soles") - Len(lsC2) + 10 - Len(Trim(pgsNomAge))) & lsC2 & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
        lsCadena = lsCadena & Space(5) & Trim(pgsNomAge) & "- " & StrConv(gcPEN_PLURAL, vbProperCase) + Space(109 - Len("- " & StrConv(gcPEN_PLURAL, vbProperCase)) - Len(lsC2) + 10 - Len(Trim(pgsNomAge))) & lsC2 & oImpresora.gPrnSaltoLinea 'MARG ERS044-2016
    Else
        lsCadena = lsCadena & Space(5) & Trim(pgsNomAge) & "- Dolares" + Space(109 - Len("- Dolares") - Len(lsC2) + 10 - Len(Trim(pgsNomAge))) & lsC2 & oImpresora.gPrnSaltoLinea
    End If
    
    lsCadena = lsCadena & CentrarCadena(psTitulo, 104) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        
    CabeceraPaginaOrdenCS = lsCadena
End Function


Private Sub cmdProcesar_Click()
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim lnSaldoIni As Double
    Dim lnIng As Double
    Dim lnSal As Double
    
    flex.Clear
    flex.Rows = 2
    flex.FormaCabecera
    
    Dim lnStockIni As Long
    Dim ldFecIni As Date
    Dim lnMontoIni As Currency
    
    If Me.txtProducto.Text = "" Then
        MsgBox "Debe ingresar el producto  del cual se desea optener el Kardex.", vbInformation, "Aviso"
        txtProducto.SetFocus
        Exit Sub
    End If
    
    If Me.txtAlmacen.Text = "" Then
        MsgBox "Debe ingresar el almacen del cual se desea optener el Kardex.", vbInformation, "Aviso"
        txtAlmacen.SetFocus
        Exit Sub
    ElseIf Not IsDate(Me.mskFecIni.Text) Then
        MsgBox "Debe ingresar una fecha correcta de inicio de proceso.", vbInformation, "Aviso"
        mskFecIni.SetFocus
        Exit Sub
    ElseIf Not IsDate(Me.mskFecFin.Text) Then
        MsgBox "Debe ingresar una fecha correcta de inicio de proceso.", vbInformation, "Aviso"
        mskFecFin.SetFocus
        Exit Sub
    
    End If
    
    ldFecIni = CDate(Me.mskFecIni.Text)
    lnStockIni = oALmacen.GetStock(Me.txtAlmacen.Text, Me.txtProducto.Text, Val(Right(Me.cboTpoAlm.Text, 5)), ldFecIni, lnMontoIni)
    
    Set rs = oALmacen.GetLogKardex(Me.txtAlmacen.Text, Me.txtProducto.Text, CDate(Me.mskFecIni.Text), CDate(Me.mskFecFin.Text), Val(Right(cboTpoAlm.Text, 5)))
    If rs.RecordCount = 0 Then
       MsgBox "No existen Registros de Almacen en el periodo seleccionado.", vbOKOnly + vbInformation, "Atención"
       Exit Sub
    End If '****NAGL 20180110
    'If rs Is Nothing Then Exit Sub 'Comentado by NAGL 20180110
    
    Me.flex.AdicionaFila
    flex.TextMatrix(flex.Rows - 1, 1) = Format(ldFecIni, gsFormatoFechaView)
    flex.TextMatrix(flex.Rows - 1, 5) = Format(lnStockIni, "#,##0.00")
    flex.TextMatrix(flex.Rows - 1, 12) = Format(lnMontoIni, "#,##0.00")
    
    lnIng = lnStockIni
    While Not rs.EOF
        Me.flex.AdicionaFila
        flex.TextMatrix(flex.Rows - 1, 1) = Format(rs!fecha, gsFormatoFechaView)
        flex.TextMatrix(flex.Rows - 1, 2) = rs!cDocNro & ""
        
        If rs!nMovEstado = 20 Then
            flex.TextMatrix(flex.Rows - 1, 3) = Format(rs!nMovCant, "#,##0.00")
            lnIng = lnIng + rs!nMovCant
            flex.TextMatrix(flex.Rows - 1, 10) = Format(rs!nMovImporte, "#,##0.00")
            If IsNumeric(flex.TextMatrix(flex.Rows - 2, 5)) Then
                flex.TextMatrix(flex.Rows - 1, 5) = Format(CCur(flex.TextMatrix(flex.Rows - 2, 5)) + CCur(flex.TextMatrix(flex.Rows - 1, 3)), "#,##0.00")
            Else
                flex.TextMatrix(flex.Rows - 1, 5) = Format(flex.TextMatrix(flex.Rows - 1, 3), "#,##0.00")
            End If
        Else
            flex.TextMatrix(flex.Rows - 1, 4) = Format(rs!nMovCant, "#,##0.00")
            lnSal = lnSal + rs!nMovCant
            flex.TextMatrix(flex.Rows - 1, 11) = Format(rs!nMovImporte, "#,##0.00")
            If IsNumeric(flex.TextMatrix(flex.Rows - 2, 5)) Then
                flex.TextMatrix(flex.Rows - 1, 5) = Format(CCur(flex.TextMatrix(flex.Rows - 2, 5)) - CCur(flex.TextMatrix(flex.Rows - 1, 4)), "#,##0.00")
            Else
                flex.TextMatrix(flex.Rows - 1, 5) = "-" & Format(flex.TextMatrix(flex.Rows - 1, 4), "#,##0.00")
            End If
        End If
        
        flex.TextMatrix(flex.Rows - 1, 6) = IIf(IsNull(rs!cPersNombre), "", rs!cPersNombre)
        flex.TextMatrix(flex.Rows - 1, 7) = IIf(IsNull(rs!cAgeCod), "", rs!cAgeCod)
        flex.TextMatrix(flex.Rows - 1, 8) = IIf(IsNull(rs!cAreaCod), "", rs!cAreaDescripcion & "")
        
        flex.TextMatrix(flex.Rows - 1, 9) = Format(IIf(IsNull(rs!PrePromedio), 0, rs!PrePromedio), "#,##0.00")
        
        If flex.TextMatrix(flex.Rows - 1, 0) = "1" Then
           If flex.TextMatrix(flex.Rows - 1, 10) <> "" Then
                flex.TextMatrix(flex.Rows - 1, 12) = Format(lnSaldoIni + CCur(flex.TextMatrix(flex.Rows - 1, 10)), "#,##0.00")
           Else
                flex.TextMatrix(flex.Rows - 1, 12) = Format(lnSaldoIni - CCur(flex.TextMatrix(flex.Rows - 1, 11)), "#,##0.00")
           End If
        Else
           If flex.TextMatrix(flex.Rows - 1, 10) <> "" Then
                flex.TextMatrix(flex.Rows - 1, 12) = Format(CCur(flex.TextMatrix(flex.Rows - 2, 12)) + CCur(flex.TextMatrix(flex.Rows - 1, 10)), "#,##0.00")
           Else
                flex.TextMatrix(flex.Rows - 1, 12) = Format(IIf(flex.TextMatrix(flex.Rows - 2, 12) = "", 0, CCur(flex.TextMatrix(flex.Rows - 2, 12))) + CCur(IIf(flex.TextMatrix(flex.Rows - 1, 11) = "", 0, flex.TextMatrix(flex.Rows - 1, 11))), "#,##0.00")
           End If
        End If
        rs.MoveNext
    Wend
    
    Me.lblSalG.Caption = Format(lnSal, "#,##0.00")
    Me.lblIngG.Caption = Format(lnIng, "#,##0.00")
    Me.lblSaldosG.Caption = Format(lnIng - lnSal, "#,##0.00")
    cmdTarjeta.Enabled = True 'FRHU 20140710 ERS048-2014
End Sub
'FRHU 20140710 ERS048-2014
Private Sub cmdTarjeta_Click()
    Dim nTotalFilas As Currency, nTotalPaginas As Currency, nFilasPorPagina As Currency
    Dim i As Currency, x As Currency, y As Currency, z As Currency, b As Currency
    On Error GoTo ErrorImprimirPDF
    Dim oDoc  As cPDF
    Set oDoc = New cPDF
    
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Administrativo"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Tarjeta Kardex " & gsCodUser
    oDoc.Title = "Tarjeta Kardex " & gsCodUser
    If Not oDoc.PDFCreate(App.path & "\Spooler\" & "TarjetaKardex" & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If
    
    'FUENTES
    Dim nFTabla As Integer
    Dim nFTablaCabecera As Integer
    Dim lnFontSizeBody As Integer
    
    oDoc.LoadImageFromFile App.path & "\Logo_NewCmacmaynas.jpg", "Logo"
    oDoc.Fonts.Add "F1", "Times New Roman", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Times New Roman", TrueType, Bold, WinAnsiEncoding
    oDoc.Fonts.Add "F3", "Arial Narrow", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F4", "Arial Narrow", TrueType, Bold, WinAnsiEncoding
    'FIN FUENTES
    nFilasPorPagina = 42
    nTotalFilas = Me.flex.Rows - 1
    'nTotalFilas = 60
    x = CCur(nTotalFilas / 42)
    'FRHU 20140917 Observacion TI-ERS091-2014
    'y = Left(x, InStr(1, CStr(x), ".") - 1)
    If InStr(1, CStr(x), ".") = 0 Then
        y = x
    Else
        y = Left(x, InStr(1, CStr(x), ".") - 1)
    End If
    'FIN FRHU
    If x > y Then y = y + 1
    If x = 0 Then MsgBox "No hay datos que mostrar": Exit Sub
    z = 1
    nTotalPaginas = 1
    Do While z <= y
    oDoc.NewPage A4_Vertical
    
    'INICIO DEL PDF: COMIENZA A ESCRIBIR EN EL PDF
    'SECCION Nº 1
    oDoc.WTextBox 70, 50, 50, 100, "", "F1", 12, hCenter, , , 1, , , 2
    oDoc.WImage 120, 51, 49, 99, "Logo"
    oDoc.WTextBox 70, 150, 50, 400, "", "F3", 7, hCenter, , , 1, vbBlack
    'CABECERA
        oDoc.WTextBox 70, 250, 50, 200, "TARJETA BIN CARD", "F2", 12, hCenter, , vbBlack, , , , 2
        oDoc.WTextBox 80, 200, 50, 300, "CONTROL FISICO / VISIBLE ALMACEN", "F2", 12, hCenter, , vbBlack, , , , 2
    'SEGUNDA CABECERA
    'Fila 1
    oDoc.WTextBox 120, 50, 13, 100, "Articulo / Descripción: ", "F2", 8, hLeft, , , 1, , , 2
    oDoc.WTextBox 120, 150, 13, 400, Mid(lblProducto.Caption, 1, InStr(1, lblProducto.Caption, "-") - 1), "F1", 8, hLeft, , , 1, , , 2
    'Fila 2
    oDoc.WTextBox 133, 50, 13, 100, "Agencia: ", "F2", 8, hLeft, , , 1, , , 2
    oDoc.WTextBox 133, 150, 13, 100, Mid(lblAlmacenG.Caption, 9, Len(lblAlmacenG.Caption)), "F1", 8, hLeft, , , 1, , , 2
    oDoc.WTextBox 133, 250, 13, 100, "Codigo: ", "F2", 8, hLeft, , , 1, , , 2
    oDoc.WTextBox 133, 350, 13, 200, txtProducto, "F1", 8, hLeft, , , 1, , , 2
    'Fila 3
    oDoc.WTextBox 146, 50, 13, 100, "Unidad Medidad: ", "F2", 8, hLeft, , , 1, , , 2
    oDoc.WTextBox 146, 150, 13, 100, Mid(lblProducto.Caption, InStr(1, lblProducto.Caption, "-") + 1, Len(lblProducto.Caption)), "F1", 8, hLeft, , , 1, , , 2
    oDoc.WTextBox 146, 250, 13, 100, "Ubicación: ", "F2", 8, hLeft, , , 1, , , 2
    oDoc.WTextBox 146, 350, 13, 30, "Est: ", "F2", 8, hCenter, , , 1, , , 2
    oDoc.WTextBox 146, 380, 13, 70, " ", "F1", 8, hCenter, , , 1, , , 2
    oDoc.WTextBox 146, 450, 13, 30, "Sec ", "F2", 8, hCenter, , , 1, , , 2
    oDoc.WTextBox 146, 480, 13, 70, " ", "F1", 8, hCenter, , , 1, , , 2
    'Fila 4
    oDoc.WTextBox 159, 50, 13, 100, "Grupo / Familia: ", "F2", 8, hLeft, , , 1, , , 2
    oDoc.WTextBox 159, 150, 13, 400, "", "F1", 8, hCenter, , , 1, , , 2
    
    'SECCION Nº 2
    oDoc.WTextBox 172, 50, 600, 500, "", "F3", 7, hCenter, , , 1, vbBlack
    'Cabecera
        oDoc.WTextBox 180, 60, 30, 100, "FECHA", "F2", 10, hCenter, , , 1, , , 2
        oDoc.WTextBox 180, 160, 30, 100, "DOCUMENTO", "F2", 10, hCenter, , , 1, , , 2
        oDoc.WTextBox 180, 260, 15, 180, "MOVIMIENTOS", "F2", 10, hCenter, , , 1, , , 2
        oDoc.WTextBox 195, 260, 15, 60, "INGRESO", "F2", 10, hCenter, , , 1, , , 2
        oDoc.WTextBox 195, 320, 15, 60, "SALIDA", "F2", 10, hCenter, , , 1, , , 2
        oDoc.WTextBox 195, 380, 15, 60, "SALDO", "F2", 10, hCenter, , , 1, , , 2
        oDoc.WTextBox 180, 440, 30, 100, "FIRMA", "F2", 10, hCenter, , , 1, , , 2
    'Cargar Datos
    Dim ntop As Integer
    ntop = 0
    For i = nTotalPaginas To nTotalFilas
        oDoc.WTextBox 210 + ntop, 60, 13, 100, Me.flex.TextMatrix(i, 1), "F1", 8, hCenter, , , 1, , , 2
        oDoc.WTextBox 210 + ntop, 160, 13, 100, Me.flex.TextMatrix(i, 2), "F1", 8, hCenter, , , 1, , , 2
        oDoc.WTextBox 210 + ntop, 260, 13, 60, Me.flex.TextMatrix(i, 3), "F1", 8, hRight, , , 1, , , 2
        oDoc.WTextBox 210 + ntop, 320, 13, 60, Me.flex.TextMatrix(i, 4), "F1", 8, hRight, , , 1, , , 2
        oDoc.WTextBox 210 + ntop, 380, 13, 60, Me.flex.TextMatrix(i, 5), "F1", 8, hRight, , , 1, , , 2
        oDoc.WTextBox 210 + ntop, 440, 13, 100, "", "F1", 8, hCenter, , , 1, , , 2
        ntop = ntop + 13
        If i = nFilasPorPagina Then Exit For
    Next i
    b = nFilasPorPagina
    nFilasPorPagina = nFilasPorPagina + 42
    nTotalPaginas = i
    nTotalPaginas = nTotalPaginas + 1
    If i <> b Then
        For i = nTotalPaginas To b
            oDoc.WTextBox 210 + ntop, 60, 13, 100, "", "F1", 8, hCenter, , , 1, , , 2
            oDoc.WTextBox 210 + ntop, 160, 13, 100, "", "F1", 8, hCenter, , , 1, , , 2
            oDoc.WTextBox 210 + ntop, 260, 13, 60, "", "F1", 8, hCenter, , , 1, , , 2
            oDoc.WTextBox 210 + ntop, 320, 13, 60, "", "F1", 8, hCenter, , , 1, , , 2
            oDoc.WTextBox 210 + ntop, 380, 13, 60, "", "F1", 8, hCenter, , , 1, , , 2
            oDoc.WTextBox 210 + ntop, 440, 13, 100, "", "F1", 8, hCenter, , , 1, , , 2
            ntop = ntop + 13
        Next i
    End If
    'FIN DEL PDF
    z = z + 1
    Loop
    oDoc.PDFClose
    oDoc.Show
    Exit Sub
ErrorImprimirPDF:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub
'FIN FRHU 20140710
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim oDoc As DOperaciones
    Set oDoc = New DOperaciones
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim oCon As DConstantes
    Set oCon = New DConstantes
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rs = oCon.GetConstante(5010, False)
    CargaCombo rs, cboTpoAlm
    cboTpoAlm.ListIndex = 0
    
    Me.txtAlmacen.rs = oDoc.GetAlmacenes
    Me.txtProducto.rs = oALmacen.GetBienesProv()
End Sub


Private Sub mskFecFin_GotFocus()
    mskFecFin.SelStart = 0
    mskFecFin.SelLength = 50
End Sub

Private Sub mskFecFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdProcesar.SetFocus
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

Private Sub txtAlmacen_EmiteDatos()
    lblAlmacenG.Caption = txtAlmacen.psDescripcion
End Sub

Private Sub txtAlmacen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.mskFecIni.SetFocus
    End If
End Sub

Private Sub txtProducto_EmiteDatos()
    lblProducto.Caption = txtProducto.psDescripcion
End Sub

Private Sub txtProducto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'Me.txtAlmacen.SetFocus
    End If
End Sub
