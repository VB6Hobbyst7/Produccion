VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCapMigracionAhorros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Migración de Cuentas de Ahorros"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15210
   Icon            =   "frmCapMigracionAhorros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   15210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FRMigracion 
      Caption         =   "Migración"
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   15015
      Begin VB.TextBox txtGlosa 
         Height          =   315
         Left            =   1680
         TabIndex        =   17
         Top             =   1080
         Width           =   8055
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   9960
         TabIndex        =   14
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   9960
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdLimpiar 
         Caption         =   "&Limpiar"
         Height          =   375
         Left            =   9960
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox cboSubProducto 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   330
         Width           =   4095
      End
      Begin SICMACT.TxtBuscar txtEmpleador 
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
         TipoBusPers     =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Glosa:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblEmpleador 
         Caption         =   "Empleador:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblEmpleador2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3960
         TabIndex        =   11
         Top             =   705
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.Label lblSubProducto 
         Caption         =   "SubProducto Destino:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame FRCuentas 
      Caption         =   "Cuenta"
      Height          =   2775
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   15015
      Begin SICMACT.FlexEdit FECuentas 
         Height          =   2415
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   4260
         Cols0           =   18
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   $"frmCapMigracionAhorros.frx":030A
         EncabezadosAnchos=   "500-1500-3000-1800-2000-1200-1200-1200-1000-0-0-1000-0-0-0-0-0-0"
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
         ColumnasAEditar =   "X-1-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-1-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-C-L-C-C-R-R-R-R-R-C-R-C-C-L-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-4-0-0-0-0-0-3-5-4-0-4"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame FRCargaArchivo 
      Caption         =   "Carga de Archivo"
      Height          =   735
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton cmdExaminar 
         Caption         =   "&Examinar"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame FRCargaManual 
      Caption         =   "Carga Manual"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "&Quitar"
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog dlgArchivo 
      Left            =   4560
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCapMigracionAhorros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'*** Nombre : frmCapMigracionAhorros
'*** Descripción : Formulario para migrar cuentas de ahorros.
'*** Creación : ELRO, 20130219 07:09:58 PM, según TI-ERS011-2013
'********************************************************************
Option Explicit

Dim fsRuta As String


Private Function ValidarCampos(Optional ByVal psCtaCod As String = "") As Boolean
ValidarCampos = False

If Trim(psCtaCod) = "" Then

    If Trim(FECuentas.TextMatrix(1, 0)) = "" Then
        MsgBox "No hay cuentas para migrar.", vbInformation, "Aviso"
        Exit Function
    End If
    
    If Trim(txtGlosa) = "" Then
        MsgBox "Debe ingresar la glosa.", vbInformation, "Aviso"
        Exit Function
    End If
    
    If CInt(Trim(Right(cboSubProducto, 3))) = 6 Then
        If Trim(txtEmpleador) = "" Then
            MsgBox "Debe ingresar el Empleador.", vbInformation, "Aviso"
             Exit Function
        End If
    End If
Else
    Dim i, j As Integer
    j = FECuentas.Rows
    
    If j > 2 Then
        For i = 1 To j - 1
            If Trim(FECuentas.TextMatrix(i, 3)) = psCtaCod Then
                MsgBox "La cuenta " & psCtaCod & " ya existe en la lista.", vbInformation, "Aviso"
                Exit Function
            End If
        Next i
    End If
    
End If

ValidarCampos = True
End Function


Private Sub CargarCombo()
Dim oNCOMCaptaGenerales As COMNCaptaGenerales.NCOMCaptaGenerales
Set oNCOMCaptaGenerales = New COMNCaptaGenerales.NCOMCaptaGenerales
Dim rsSubProductoDestino As ADODB.Recordset
Set rsSubProductoDestino = New ADODB.Recordset

Set rsSubProductoDestino = oNCOMCaptaGenerales.obtenerSubProductoDestino

Do While Not rsSubProductoDestino.EOF
    cboSubProducto.AddItem rsSubProductoDestino!cDescripcionSubProducto & Space(100) & rsSubProductoDestino!nCodigoSubProducto
    rsSubProductoDestino.MoveNext
Loop

cboSubProducto.ListIndex = 0

End Sub

Private Function cargarTEA(ByVal pnProducto As Integer, ByVal pnMoneda As Integer, _
                           ByVal pnTipoTasa As Integer, ByVal pnPlazo As Long, _
                           ByVal pnMonto As Double, ByVal pbOrdPag As Boolean, _
                           ByVal pSubPrograma As Integer) As Double
                           
Dim oNCOMCaptaDefinicion As COMNCaptaGenerales.NCOMCaptaDefinicion
Set oNCOMCaptaDefinicion = New COMNCaptaGenerales.NCOMCaptaDefinicion

cargarTEA = oNCOMCaptaDefinicion.GetCapTasaInteres(pnProducto, pnMoneda, pnTipoTasa, pnPlazo, pnMonto, gsCodAge, pbOrdPag, pSubPrograma)
End Function

Private Sub limpiarCampos()
Call LimpiaFlex(FECuentas)
cboSubProducto.ListIndex = 0
txtEmpleador.Visible = False
txtEmpleador = ""
lblEmpleador.Visible = False
lblEmpleador2.Visible = False
lblEmpleador2 = ""
txtGlosa = ""
End Sub

Private Sub cargarArchivo()

If Trim(fsRuta) = "" Then Exit Sub

Dim oNCOMCaptaGenerales As COMNCaptaGenerales.NCOMCaptaGenerales
Set oNCOMCaptaGenerales = New COMNCaptaGenerales.NCOMCaptaGenerales

Dim rsCuenta As ADODB.Recordset
Set rsCuenta = New ADODB.Recordset

Dim rsTitular As ADODB.Recordset
Set rsTitular = New ADODB.Recordset
Dim lsCuenta As String

Dim lbError As Boolean

Dim fs As Scripting.FileSystemObject

lbError = False
If InStr(Trim(UCase(fsRuta)), ".XLS") <> 0 Or InStr(Trim(UCase(fsRuta)), ".XLSX") <> 0 Then

    'Variable de tipo Aplicación de Excel
    Dim oExcel As Excel.Application
    Dim lnTipoDOI, lnFila1, lnFila2, lnFilasFormato As Integer
    Dim lsDOI As String
    Dim lsMoneda As String
    Dim lbExisteCTS As Boolean
    Dim lbExisteError As Boolean
    
    '***Para verificar la existencia del archivo en la ruta
    Set fs = New Scripting.FileSystemObject
    

  
    'Una variable de tipo Libro de Excel
    Dim oLibro As Excel.Workbook
    Dim oHoja As Excel.Worksheet

    'creamos un nuevo objeto excel
    Set oExcel = New Excel.Application
     
    lnFilasFormato = 3525
 
    'Usamos el método open para abrir el archivo que está en el directorio del programa llamado archivo.xls
    
    If fs.FileExists(fsRuta) Then
        Set oLibro = oExcel.Workbooks.Open(fsRuta)
    Else
        MsgBox "No existe el archivo en esta ruta: " & fsRuta, vbCritical, "Advertencia"
        Set oHoja = Nothing
        Set oLibro = Nothing
        Set oExcel = Nothing
        Set fs = Nothing
        Exit Sub
    End If
    
    'Hacemos referencia a la Hoja
    Set oHoja = oLibro.Sheets(1)
    
    'Hacemos el Excel Visible
    'oLibro.Visible = False


    lsCuenta = ""
    With oHoja
        For lnFila1 = 2 To lnFilasFormato
            lsCuenta = .Cells(lnFila1, 2)
        
            If Len(lsCuenta) = 18 Then
                If ValidarCampos(lsCuenta) Then
                    Set rsTitular = oNCOMCaptaGenerales.ObtenerTitularCuenta(lsCuenta)
                    Set rsCuenta = oNCOMCaptaGenerales.obtenerDatosCuentaPersona(lsCuenta)
                    If Not (rsCuenta.BOF And rsCuenta.EOF) Then
                        FECuentas.lbEditarFlex = True
                        FECuentas.SetFocus
                        FECuentas.AdicionaFila
                        FECuentas.TextMatrix(FECuentas.row, 1) = rsTitular!codigo
                        FECuentas.TextMatrix(FECuentas.row, 2) = rsTitular!cPersNombre
                        FECuentas.TextMatrix(FECuentas.row, 3) = lsCuenta
                        FECuentas.TextMatrix(FECuentas.row, 4) = rsCuenta!cSubProducto
                        FECuentas.TextMatrix(FECuentas.row, 5) = rsCuenta!cFechaApertura
                        FECuentas.TextMatrix(FECuentas.row, 6) = rsCuenta!cMonena
                        FECuentas.TextMatrix(FECuentas.row, 7) = Format$(rsCuenta!nSaldo, "##,##0.00")
                        FECuentas.TextMatrix(FECuentas.row, 8) = Format$(rsCuenta!TEA, "#,##0.00")
                        FECuentas.TextMatrix(FECuentas.row, 9) = rsCuenta!nTasaNominalOrigen
                        FECuentas.TextMatrix(FECuentas.row, 10) = cargarTEA(232, rsCuenta!nmoneda, 100, 0, rsCuenta!nSaldo, 0, CInt(Trim(Right(cboSubProducto, 3))))
                        FECuentas.TextMatrix(FECuentas.row, 11) = Format$(ConvierteTNAaTEA(CDbl(FECuentas.TextMatrix(FECuentas.row, 10))), "#,##0.00")
                        FECuentas.TextMatrix(FECuentas.row, 12) = "L"
                        FECuentas.TextMatrix(FECuentas.row, 13) = rsCuenta!nSubProducto
                        FECuentas.TextMatrix(FECuentas.row, 14) = rsCuenta!dUltCierre
                        FECuentas.TextMatrix(FECuentas.row, 15) = rsCuenta!nSaldoDisp
                        FECuentas.TextMatrix(FECuentas.row, 16) = rsCuenta!cTipoCuenta
                        FECuentas.TextMatrix(FECuentas.row, 17) = rsCuenta!nIntAcum
                        FECuentas.lbEditarFlex = False
                    Else
                         frmCapMigracionAhorrosError.CargarDatos lsCuenta, "Esta cuenta no existe o es un subprograma que no esta contemplado para la migración.", lbError
                         lbError = True
                    End If
                Else
                    FECuentas.lbEditarFlex = False
                    FECuentas.EliminaFila FECuentas.row
                End If
            ElseIf Len(lsCuenta) < 18 And Len(lsCuenta) > 0 Then
                frmCapMigracionAhorrosError.CargarDatos lsCuenta, "Esta cuenta no tiene formato.", lbError
                lbError = True
            End If
        Next lnFila1
        
        If lbError Then
            frmCapMigracionAhorrosError.Show 1
            Call LimpiaFlex(FECuentas)
        End If
        
        Set rsTitular = Nothing
        Set rsCuenta = Nothing
    End With
    
oLibro.Close
oExcel.Quit
Set oHoja = Nothing
Set oLibro = Nothing
Set oExcel = Nothing

ElseIf InStr(Trim(UCase(fsRuta)), ".TXT") <> 0 Then
    Dim f As Integer
    Dim str_Linea As String
    
    '***Para verificar la existencia del archivo en la ruta
    Set fs = New Scripting.FileSystemObject
    
    If Not fs.FileExists(fsRuta) Then
        MsgBox "No existe el archivo en esta ruta: " & fsRuta, vbCritical, "Advertencia"
        Set rsTitular = Nothing
        Set rsCuenta = Nothing
        Set fs = Nothing
        Exit Sub
    End If
        
    f = FreeFile
    lsCuenta = ""
    Open fsRuta For Input As #f
        'Inserta Detalle de Recaudo Temporal
        Do
            Line Input #f, str_Linea
            If Len(str_Linea) = 18 Then
                lsCuenta = Left(str_Linea, 18)
                
                Set rsTitular = oNCOMCaptaGenerales.ObtenerTitularCuenta(lsCuenta)
                Set rsCuenta = oNCOMCaptaGenerales.obtenerDatosCuentaPersona(lsCuenta)
                If Not (rsCuenta.BOF And rsCuenta.EOF) Then
                    FECuentas.lbEditarFlex = True
                    FECuentas.SetFocus
                    FECuentas.AdicionaFila
                    FECuentas.TextMatrix(FECuentas.row, 1) = rsTitular!codigo
                    FECuentas.TextMatrix(FECuentas.row, 2) = rsTitular!cPersNombre
                    FECuentas.TextMatrix(FECuentas.row, 3) = Left(str_Linea, 18)
                    FECuentas.TextMatrix(FECuentas.row, 4) = rsCuenta!cSubProducto
                    FECuentas.TextMatrix(FECuentas.row, 5) = rsCuenta!cFechaApertura
                    FECuentas.TextMatrix(FECuentas.row, 6) = rsCuenta!cMonena
                    FECuentas.TextMatrix(FECuentas.row, 7) = Format$(rsCuenta!nSaldo, "##,##0.00")
                    FECuentas.TextMatrix(FECuentas.row, 8) = Format$(rsCuenta!TEA, "#,##0.00")
                    FECuentas.TextMatrix(FECuentas.row, 9) = rsCuenta!nTasaNominalOrigen
                    FECuentas.TextMatrix(FECuentas.row, 10) = cargarTEA(232, rsCuenta!nmoneda, 100, 0, rsCuenta!nSaldo, 0, CInt(Trim(Right(cboSubProducto, 3))))
                    FECuentas.TextMatrix(FECuentas.row, 11) = Format$(ConvierteTNAaTEA(CDbl(FECuentas.TextMatrix(FECuentas.row, 10))), "#,##0.00")
                    FECuentas.TextMatrix(FECuentas.row, 12) = "L"
                    FECuentas.TextMatrix(FECuentas.row, 13) = rsCuenta!nSubProducto
                    FECuentas.TextMatrix(FECuentas.row, 14) = rsCuenta!dUltCierre
                    FECuentas.TextMatrix(FECuentas.row, 15) = rsCuenta!nSaldoDisp
                    FECuentas.TextMatrix(FECuentas.row, 16) = rsCuenta!cTipoCuenta
                    FECuentas.TextMatrix(FECuentas.row, 17) = rsCuenta!nIntAcum
                    FECuentas.lbEditarFlex = False
                Else
                    frmCapMigracionAhorrosError.CargarDatos lsCuenta, "Esta cuenta no existe o es un subprograma que no esta contemplado para la migración.", lbError
                    lbError = True
                End If
            ElseIf Len(lsCuenta) < 18 And Len(lsCuenta) > 0 Then
                frmCapMigracionAhorrosError.CargarDatos lsCuenta, "Esta cuenta no tiene formato.", lbError
            End If
        Loop While Not EOF(f)
    Close #f
    If lbError Then
        frmCapMigracionAhorrosError.Show 1
        Call LimpiaFlex(FECuentas)
    End If
    Set rsTitular = Nothing
    Set rsCuenta = Nothing
End If
Set fs = Nothing
End Sub

Private Sub cboSubProducto_Click()
If Trim(Right(cboSubProducto, 3)) = "6" Then '6:Caja Sueldo
    txtEmpleador.Visible = True
    lblEmpleador.Visible = True
    lblEmpleador2.Visible = True
Else
    txtEmpleador.Visible = False
    txtEmpleador = ""
    lblEmpleador.Visible = False
    lblEmpleador2.Visible = False
    lblEmpleador2 = ""
End If

If Trim(FECuentas.TextMatrix(1, 0)) <> "" Then
    Dim i, j As Integer
    j = FECuentas.Rows
    i = 0
    For i = 1 To j - 1
        FECuentas.TextMatrix(i, 10) = cargarTEA(232, IIf(FECuentas.TextMatrix(i, 6) = "SOLES", 1, 2), 100, 0, FECuentas.TextMatrix(i, 7), 0, CInt(Trim(Right(cboSubProducto, 3))))
        FECuentas.TextMatrix(i, 11) = Format$(ConvierteTNAaTEA(CDbl(FECuentas.TextMatrix(i, 10))), "#,##0.00")
    Next i
End If
End Sub

Private Sub cmdAgregar_Click()
If Trim(FECuentas.TextMatrix(1, 0)) = "" Then
    FECuentas.AdicionaFila
    FECuentas.SetFocus
    FECuentas.lbEditarFlex = True
Else
    If FECuentas.TextMatrix(FECuentas.row, 12) = "L" Then
        If MsgBox("¿Se eliminará los registros previos antes de agregar un nuevo registro?", vbYesNo, "Aviso") = vbYes Then
            Call LimpiaFlex(FECuentas)
            FECuentas.AdicionaFila
            FECuentas.SetFocus
            FECuentas.lbEditarFlex = True
        End If
    Else
        FECuentas.AdicionaFila
        FECuentas.SetFocus
        FECuentas.lbEditarFlex = True
    End If
End If
End Sub

Private Sub cmdExaminar_Click()

If Trim(FECuentas.TextMatrix(1, 0)) <> "" Then
    If MsgBox("¿Se eliminará los registros previos antes de cargar los datos del archivo?", vbYesNo, "Aviso") = vbYes Then
        Call LimpiaFlex(FECuentas)
        fsRuta = Empty
        dlgArchivo.InitDir = "C:\"
        dlgArchivo.Filter = "Archivos de Texto (*.txt)|*.txt|Archivos de Excel (*.xls)|*.xls| Archivos de Excel (*.xlsx)|*.xlsx"
        dlgArchivo.ShowOpen
        If dlgArchivo.FileName <> Empty Then
            fsRuta = dlgArchivo.FileName
        Else
            MsgBox "No se eligio un archivo"
            Exit Sub
        End If
        FRCargaManual.Enabled = False
        FRMigracion.Enabled = False
        cargarArchivo
    End If
Else
    fsRuta = Empty
    dlgArchivo.InitDir = "C:\"
    dlgArchivo.Filter = "Archivos de Texto (*.txt)|*.txt|Archivos de Excel (*.xls)|*.xls| Archivos de Excel (*.xlsx)|*.xlsx"
    dlgArchivo.ShowOpen
    If dlgArchivo.FileName <> Empty Then
        fsRuta = dlgArchivo.FileName
    Else
        MsgBox "No se eligio un archivo"
        Exit Sub
    End If
    FRCargaManual.Enabled = False
    FRMigracion.Enabled = False
    cargarArchivo
End If
FRCargaManual.Enabled = True
FRMigracion.Enabled = True
End Sub

Private Sub cmdGrabar_Click()

If ValidarCampos = False Then Exit Sub
If MsgBox("¿Está seguro que desea migrar al SubProducto " & Trim(Left(cboSubProducto, 50)) & "?", vbYesNo, "Aviso") = vbYes Then
'If MsgBox("¿Esta seguro que desea migrar al SubProducto " & Left(cboSubProducto, 50) & "?", vbYesNo, "Aviso") = vbYes Then ' COMENTADO POR PTI1 24092018
        
    Dim oNCOMCaptaMovimiento As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set oNCOMCaptaMovimiento = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim oNCOMContFunciones As COMNContabilidad.NCOMContFunciones
    Set oNCOMContFunciones = New COMNContabilidad.NCOMContFunciones
    Dim oclsprevio As previo.clsprevio
    Set oclsprevio = New previo.clsprevio
    
    Dim rsCuentas As ADODB.Recordset
    Set rsCuentas = New ADODB.Recordset
    Dim lsMovNro As String
    Dim lnConfirmar As Long
    Dim lsImpFirmas As String
    Dim lsImpBoleta As String
    
    FRCargaManual.Enabled = False
    FRCargaArchivo.Enabled = False
    FRMigracion.Enabled = False
    
    Set rsCuentas = FECuentas.GetRsNew()
    lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    '***Migracion de las cuentas****
    lnConfirmar = oNCOMCaptaMovimiento.migrarCuentasAhorros(lsMovNro, gAhoMigracion, txtGlosa, rsCuentas, CInt(Trim(Right(cboSubProducto, 3))), txtEmpleador, gsNomAge, lsImpFirmas)
    '***Fin Migracion de las cuentas
    
    If lnConfirmar > 0 Then
        
        '***Documentos de la apertura****
        If rsCuentas.RecordCount = 1 Then
            MsgBox "Coloque papel para el Registro de Firmas.", vbInformation, "Aviso"
            oclsprevio.Show lsImpFirmas, "Migración de Cuenta(s)", True, , gImpresora
            Set oclsprevio = Nothing
            
            MsgBox "Coloque papel para Cartillas", vbInformation, "Aviso"
            ImpreCartillaAHLote2 rsCuentas, CInt(Trim(Right(cboSubProducto, 3))), , lsMovNro
        Else
            If MsgBox("¿Desea imprimir los documentos de aperturas?", vbYesNo, "Aviso") = vbYes Then
                MsgBox "Coloque papel para el Registro de Firmas.", vbInformation, "Aviso"
                oclsprevio.Show lsImpFirmas, "Migración de Cuenta(s)", True, , gImpresora
                Set oclsprevio = Nothing
            
                MsgBox "Coloque papel para Cartillas", vbInformation, "Aviso"
                ImpreCartillaAHLote2 rsCuentas, CInt(Trim(Right(cboSubProducto, 3))), , lsMovNro
            End If
        End If
        '***Fin Documentos de la apertura
        
        If MsgBox("¿Desea realizar otra migración?", vbYesNo, "Aviso") = vbYes Then
            limpiarCampos
        Else
            limpiarCampos
            Set oNCOMCaptaMovimiento = Nothing
            Set oNCOMContFunciones = Nothing
            Unload Me
        End If
    Else
        MsgBox "No se realizo la migración de la(s) cuenta(s).", vbCritical, "Aviso"
    End If
    
    Set oNCOMCaptaMovimiento = Nothing
    Set oNCOMContFunciones = Nothing
    
     FRCargaManual.Enabled = True
    FRCargaArchivo.Enabled = True
    FRMigracion.Enabled = True
    
End If
End Sub

Private Sub cmdLimpiar_Click()
limpiarCampos
End Sub

Private Sub cmdQuitar_Click()

If Trim(FECuentas.TextMatrix(1, 0)) = "" Then
    MsgBox "No existe registro para eliminar.", vbInformation, "Aviso"
    cmdAgregar.SetFocus
    Exit Sub
Else
    FECuentas.EliminaFila (FECuentas.row)
    cmdAgregar.SetFocus
End If
End Sub

Private Sub cmdsalir_Click()
limpiarCampos
Unload Me
End Sub

Private Sub FECuentas_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
Dim oNCOMCaptaGenerales As COMNCaptaGenerales.NCOMCaptaGenerales
Set oNCOMCaptaGenerales = New COMNCaptaGenerales.NCOMCaptaGenerales
Dim rsCuentas As ADODB.Recordset
Set rsCuentas = New ADODB.Recordset
Dim rsCuenta As ADODB.Recordset
Set rsCuenta = New ADODB.Recordset

If FECuentas.PersPersoneria <> gPersonaNat Then
    MsgBox "Debe ingresar Persona Natural.", vbInformation, "Aviso"
    FECuentas.TextMatrix(FECuentas.row, 1) = ""
    FECuentas.TextMatrix(FECuentas.row, 2) = ""
    Set oNCOMCaptaGenerales = Nothing
    Set rsCuentas = Nothing
    Set rsCuenta = Nothing
    Exit Sub
End If

Dim oUCapCuenta As UCapCuenta

Set rsCuentas = oNCOMCaptaGenerales.obtenerCuentasPersona(psDataCod)

If Not (rsCuentas.EOF And rsCuentas.EOF) Then
    Do While Not rsCuentas.EOF
        frmCapMantenimientoCtas.lstCuentas.AddItem rsCuentas("cCtaCod") & Space(2) & rsCuentas("cRelacion") & Space(2) & Trim(rsCuentas("cEstado"))
        rsCuentas.MoveNext
    Loop
        
    Set oUCapCuenta = New UCapCuenta
    Set oUCapCuenta = frmCapMantenimientoCtas.Inicia
        
    If Not oUCapCuenta Is Nothing Then
        If oUCapCuenta.sCtaCod <> "" Then
            If ValidarCampos(oUCapCuenta.sCtaCod) Then
                Set rsCuenta = oNCOMCaptaGenerales.obtenerDatosCuentaPersona(oUCapCuenta.sCtaCod)
                FECuentas.TextMatrix(FECuentas.row, 3) = oUCapCuenta.sCtaCod
                FECuentas.TextMatrix(FECuentas.row, 4) = rsCuenta!cSubProducto
                FECuentas.TextMatrix(FECuentas.row, 5) = rsCuenta!cFechaApertura
                FECuentas.TextMatrix(FECuentas.row, 6) = rsCuenta!cMonena
                FECuentas.TextMatrix(FECuentas.row, 7) = Format$(rsCuenta!nSaldo, "##,##0.00")
                FECuentas.TextMatrix(FECuentas.row, 8) = Format$(rsCuenta!TEA, "#,##0.00")
                FECuentas.TextMatrix(FECuentas.row, 9) = rsCuenta!nTasaNominalOrigen
                FECuentas.TextMatrix(FECuentas.row, 10) = cargarTEA(232, rsCuenta!nmoneda, 100, 0, rsCuenta!nSaldo, 0, CInt(Trim(Right(cboSubProducto, 3))))
                FECuentas.TextMatrix(FECuentas.row, 11) = Format$(ConvierteTNAaTEA(CDbl(FECuentas.TextMatrix(FECuentas.row, 10))), "#,##0.00")
                FECuentas.TextMatrix(FECuentas.row, 12) = "M"
                FECuentas.TextMatrix(FECuentas.row, 13) = rsCuenta!nSubProducto
                FECuentas.TextMatrix(FECuentas.row, 14) = rsCuenta!dUltCierre
                FECuentas.TextMatrix(FECuentas.row, 15) = rsCuenta!nSaldoDisp
                FECuentas.TextMatrix(FECuentas.row, 16) = rsCuenta!cTipoCuenta
                FECuentas.TextMatrix(FECuentas.row, 17) = rsCuenta!nIntAcum
            Else
                    FECuentas.TextMatrix(FECuentas.row, 1) = ""
                    FECuentas.TextMatrix(FECuentas.row, 2) = ""
                    Set oNCOMCaptaGenerales = Nothing
                    Set oUCapCuenta = Nothing
                    Set rsCuentas = Nothing
                    Set rsCuenta = Nothing
                    Exit Sub
            End If
        Else
            FECuentas.TextMatrix(FECuentas.row, 1) = ""
            FECuentas.TextMatrix(FECuentas.row, 2) = ""
            Set oNCOMCaptaGenerales = Nothing
            Set oUCapCuenta = Nothing
            Set rsCuentas = Nothing
            Set rsCuenta = Nothing
            Exit Sub
        End If
    End If
Else
    MsgBox "Persona no posee ninguna cuenta de captaciones.", vbInformation, "Aviso"
    FECuentas.TextMatrix(FECuentas.row, 1) = ""
    FECuentas.TextMatrix(FECuentas.row, 2) = ""
    Set oUCapCuenta = Nothing
    Set rsCuentas = Nothing
    Set rsCuenta = Nothing
    Exit Sub
End If

Set oUCapCuenta = Nothing
Set rsCuentas = Nothing
Set rsCuenta = Nothing
FECuentas.lbEditarFlex = False
cmdAgregar.SetFocus

End Sub

Private Sub Form_Load()
CargarCombo
End Sub


Private Sub txtEmpleador_EmiteDatos()
lblEmpleador2 = txtEmpleador.psDescripcion
txtGlosa.SetFocus
End Sub

Private Sub txtGlosa_LostFocus()
cmdGrabar.SetFocus
End Sub
