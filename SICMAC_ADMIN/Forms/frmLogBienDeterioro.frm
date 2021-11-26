VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogBienDeterioro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Deterioro de Activos Fijos"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12225
   Icon            =   "frmLogBienDeterioro.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   12225
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Información de Deterioro"
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
      Height          =   735
      Left            =   60
      TabIndex        =   11
      Top             =   5745
      Width           =   12135
      Begin VB.TextBox txtGlosa 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3480
         MaxLength       =   300
         TabIndex        =   4
         Top             =   240
         Width           =   8505
      End
      Begin MSComCtl2.DTPicker txtFecha 
         Height          =   315
         Left            =   720
         TabIndex        =   3
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   63176705
         CurrentDate     =   41414
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha:"
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   285
         Width           =   570
      End
      Begin VB.Label lblComentario 
         Caption         =   "Glosa:"
         Height          =   210
         Left            =   2520
         TabIndex        =   12
         Top             =   285
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "&Exportar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8730
      TabIndex        =   5
      Top             =   6525
      Width           =   1125
   End
   Begin VB.Frame fraOpe 
      Caption         =   "Activo Fijo"
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
      Height          =   1400
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   12120
      Begin VB.TextBox txtBSNombre 
         Appearance      =   0  'Flat
         Height          =   320
         Left            =   3690
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   600
         Width           =   4860
      End
      Begin VB.TextBox txtAreaAgeNombre 
         Appearance      =   0  'Flat
         Height          =   320
         Left            =   3690
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   4860
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10990
         TabIndex        =   17
         Top             =   960
         Width           =   1050
      End
      Begin VB.TextBox txtSerieNombre 
         Appearance      =   0  'Flat
         Height          =   320
         Left            =   3690
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   960
         Width           =   4860
      End
      Begin Sicmact.TxtBuscar txtBSCod 
         Height          =   345
         Left            =   1440
         TabIndex        =   1
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   609
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
      Begin Sicmact.TxtBuscar txtAreaAgeCod 
         Height          =   345
         Left            =   1440
         TabIndex        =   0
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   609
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
      Begin Sicmact.TxtBuscar txtSerieCod 
         Height          =   345
         Left            =   1440
         TabIndex        =   15
         Top             =   960
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   609
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Serie :"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   990
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Área/Agencia :"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   270
         Width           =   1080
      End
      Begin VB.Label lblBien 
         AutoSize        =   -1  'True
         Caption         =   "Categoría :"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11050
      TabIndex        =   7
      Top             =   6525
      Width           =   1125
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9895
      TabIndex        =   6
      Top             =   6525
      Width           =   1125
   End
   Begin Sicmact.FlexEdit FeAdj 
      Height          =   4095
      Left            =   60
      TabIndex        =   2
      Top             =   1575
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   7223
      Cols0           =   17
      HighLight       =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      EncabezadosNombres=   "Nº-OK-Serie-Descripcion-Fecha-Valor Ini.-Valor Depre.-Por Depre.-Area-Agencia-nMovNro-nAnio-cBSCod-cBaja-Moneda-Ban-ValorResidual"
      EncabezadosAnchos=   "400-400-1800-3500-1200-1200-1200-1200-800-800-0-0-1200-0-0-0-0"
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-1-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-4-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-L-R-R-R-R-C-C-C-C-C-L-C-C-C"
      FormatosEdit    =   "0-0-0-0-5-2-2-2-0-0-0-0-0-0-0-0-0"
      AvanceCeldas    =   1
      TextArray0      =   "Nº"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmLogBienDeterioro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************
'** Nombre : frmLogBienDeterioro
'** Descripción : Registro de Deterioro de los Biene creado segun ERS059-2013
'** Creación : EJVG, 20130701 09:00:00 AM
'*****************************************************************************
Option Explicit

Private Sub Form_Load()
    CentraForm Me
    CargaTxt
    Limpiar
End Sub
Private Sub CargaTxt()
    Dim oBien As New DBien
    Dim oArea As New DActualizaDatosArea
    txtAreaAgeCod.rs = oArea.GetAgenciasAreas()
    txtBSCod.rs = oBien.RecuperaCategoriasBienPaObjeto(True, "")
    txtSerieCod.rs = oBien.RecuperaSeriesPaObjeto("", "", True)
    Set oBien = Nothing
    Set oArea = Nothing
End Sub
Private Sub txtAreaAgeCod_EmiteDatos()
    Dim oBien As New DBien
    Dim lsAreaAgeCod As String
    
    Screen.MousePointer = 11
    txtAreaAgeNombre.Text = ""
    txtBSCod.Text = ""
    txtBSNombre.Text = ""
    If txtAreaAgeCod.Text <> "" Then
        txtAreaAgeNombre.Text = txtAreaAgeCod.psDescripcion
        lsAreaAgeCod = Left(Me.txtAreaAgeCod.Text, 3) & IIf(Mid(txtAreaAgeCod.Text, 4, 2) = "", "01", Mid(txtAreaAgeCod.Text, 4, 2))
        txtBSCod.rs = oBien.RecuperaCategoriasBienPaObjeto(False, lsAreaAgeCod)
    Else
        txtBSCod.rs = oBien.RecuperaCategoriasBienPaObjeto(True, "")
    End If
    txtBSCod_EmiteDatos
    Screen.MousePointer = 0
    Set oBien = Nothing
End Sub
Private Sub txtAreaAgeCod_LostFocus()
    If txtAreaAgeCod.Text = "" Then
        txtAreaAgeNombre.Text = ""
    End If
End Sub
Private Sub txtBSCod_EmiteDatos()
    Dim oBien As New DBien
    Dim lsAreaAgeCod As String
    
    Screen.MousePointer = 11
    txtBSNombre.Text = ""
    If txtBSCod.Text <> "" Then
        txtBSNombre.Text = txtBSCod.psDescripcion
    End If
    If txtAreaAgeCod.Text <> "" Then
        lsAreaAgeCod = Left(txtAreaAgeCod.Text, 3) & IIf(Mid(txtAreaAgeCod.Text, 4, 2) = "", "01", Mid(txtAreaAgeCod.Text, 4, 2))
    End If
    txtSerieCod.Text = ""
    txtSerieNombre.Text = ""
    txtSerieCod.rs = oBien.RecuperaSeriesPaObjeto(lsAreaAgeCod, txtBSCod.Text, True)
    txtSerieCod_EmiteDatos
    Screen.MousePointer = 0
    Set oBien = Nothing
End Sub
Private Sub txtBSCod_LostFocus()
    If txtBSCod.Text = "" Then
        txtBSNombre.Text = ""
    End If
End Sub
Private Sub txtSerieCod_EmiteDatos()
    txtSerieNombre.Text = ""
    If txtSerieCod.Text <> "" Then
        txtSerieNombre.Text = txtSerieCod.psDescripcion
    End If
    Call LimpiaFlex(FeAdj)
End Sub
Private Sub txtSerieCod_LostFocus()
    If txtSerieCod.Text = "" Then
        txtSerieNombre.Text = ""
    End If
End Sub
Private Sub cmdBuscar_Click()
    Dim oBien As New DBien
    Dim rsSeries As New ADODB.Recordset
    Dim lsAreaAgeCod As String
    
    Screen.MousePointer = 11
    If Me.txtAreaAgeCod.Text <> "" Then
        lsAreaAgeCod = Left(txtAreaAgeCod.Text, 3) & IIf(Mid(txtAreaAgeCod.Text, 4, 2) = "", "01", Mid(txtAreaAgeCod.Text, 4, 2))
    End If
    Set rsSeries = oBien.GetAFBienesPaDeterioro(lsAreaAgeCod, txtBSCod.Text, txtSerieCod.Text)
    Call LimpiaFlex(FeAdj)
    If rsSeries.RecordCount = 0 Then
        Screen.MousePointer = 0
        MsgBox "No se encontraron resultados de la Búsqueda realizada", vbInformation, "Aviso"
        Exit Sub
    End If
    FeAdj.rsFlex = rsSeries
    FeAdj.TopRow = 1
    FeAdj.Row = 1
    Screen.MousePointer = 0
    Set oBien = Nothing
    Set rsSeries = Nothing
End Sub
Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtGlosa.SetFocus
    End If
End Sub
Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGrabar.SetFocus
    End If
End Sub
Private Sub Limpiar()
    txtFecha.value = Format(gdFecSis, gsFormatoFechaView)
    txtGlosa.Text = ""
End Sub
Private Sub cmdExportar_Click()
    Dim xlsAplicacion As New Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlsHoja As Excel.Worksheet
    Dim lnFila As Long, lnColumna As Long, lnColumnaMax As Long
    Dim i As Long, J As Long
    Dim lsArchivo As String
    Dim bOK As Boolean
    
On Error GoTo ErrExportar
    
    If FlexVacio(FeAdj) Then
        MsgBox "No hay información para exportar a formato Excel", vbInformation, "Aviso"
        Exit Sub
    Else 'Se haya seleccionado registros
        For i = 0 To FeAdj.Rows - 1
            If FeAdj.TextMatrix(i, 1) = "." Then 'OK
                bOK = True
                Exit For
            End If
        Next
        If Not bOK Then
            MsgBox "No hay información para exportar a formato Excel", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    
    Screen.MousePointer = 11
    
    lsArchivo = "\spooler\RptDeterioroBien" & UCase(gsCodUser) & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    Set xlsLibro = xlsAplicacion.Workbooks.Add

    Set xlsHoja = xlsLibro.Worksheets.Add
    xlsHoja.Name = "Reporte Deterioro de Bien"
    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 9
    
    lnFila = 2
    
    For i = 0 To FeAdj.Rows - 1
        lnColumna = 2
        If i = 0 Or (i > 0 And FeAdj.TextMatrix(i, 1) = ".") Then 'OK
            For J = 0 To FeAdj.Cols - 1
                If J > 1 And FeAdj.ColWidth(J) > 0 Then
                    xlsHoja.Cells(lnFila, lnColumna) = FeAdj.TextMatrix(i, J)
                    lnColumna = lnColumna + 1
                    lnColumnaMax = lnColumna
                End If
            Next
            lnFila = lnFila + 1
        End If
    Next

    xlsHoja.Range(xlsHoja.Cells(2, 2), xlsHoja.Cells(2, lnColumnaMax - 1)).Interior.Color = RGB(191, 191, 191)
    xlsHoja.Range(xlsHoja.Cells(2, 2), xlsHoja.Cells(2, lnColumnaMax - 1)).Font.Bold = True
    xlsHoja.Range(xlsHoja.Cells(2, 2), xlsHoja.Cells(2, lnColumnaMax - 1)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(2, 2), xlsHoja.Cells(lnFila - 1, lnColumnaMax - 1)).Borders.Weight = xlThin

    xlsHoja.Range(xlsHoja.Cells(2, 2), xlsHoja.Cells(lnFila - 1, lnColumnaMax - 1)).EntireColumn.AutoFit
    
    MsgBox "Se ha exportado satisfactoriamente la información", vbInformation, "Aviso"
    
    xlsHoja.SaveAs App.path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    
    Screen.MousePointer = 0
    
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlsHoja = Nothing
    Exit Sub
ErrExportar:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdGrabar_Click()
    Dim oMov As DMov, bTransMov As Boolean
    Dim oPrevio As clsPrevio
    Dim oAsiento As NContImprimir
    Dim oFun As NContFunciones
    Dim Movs As Variant

    Dim i As Long
    Dim lbOK As Boolean
    Dim lsValFecha As String
    
    Dim lnMovNroAF As Long
    Dim ldFecha As Date
    Dim lsGlosa As String
    Dim lsMovNro As String, lsMovNroTotal As String
    Dim lnMovNro As Long
    Dim lnMovItem As Integer
    Dim lsSerie As String, lsBSCod As String, lsAgeCod As String
    Dim lnBANCod As Integer
    Dim lnMontoxDepreciar As Currency
    Dim lnValorResidual As Integer
    Dim lsPlantilla1 As String, lsPlantilla2 As String
    Dim lsCtaCont As String
    Dim MatSoles As Variant, MatDolares As Variant, MatDatos As Variant
    Dim index As Long
    Dim iMoneda As Integer, iSleep As Long
    Dim lnNroReg As Integer
    
    On Error GoTo ErrGrabar
    
    If FlexVacio(FeAdj) Then
        MsgBox "Ud. debe de seleccionar los registros que se van a registrar su Deterioro", vbInformation, "Aviso"
        Exit Sub
    Else
        For i = 1 To FeAdj.Rows - 1
            If FeAdj.TextMatrix(i, 1) = "." And CCur(Trim(FeAdj.TextMatrix(i, 7))) > 0 Then
                lbOK = True
                Exit For
            End If
        Next
        If Not lbOK Then
            MsgBox "Ud. debe de seleccionar los registros que se van a registrar su Deterioro," & Chr(10) & "y verificar que el monto por depreciar sea mayor a cero", vbInformation, "Aviso"
            Exit Sub
        End If
        lsValFecha = ValidaFecha(txtFecha.value)
        If lsValFecha <> "" Then
            MsgBox lsValFecha, vbInformation, "Aviso"
            txtFecha.SetFocus
            Exit Sub
        End If
        If Len(Trim(txtGlosa.Text)) = 0 Then
            MsgBox "Ud. de ingresar la Glosa del Deterioro", vbInformation, "Aviso"
            txtGlosa.SetFocus
            Exit Sub
        End If
    End If
    
    ldFecha = CDate(txtFecha.value)
    Set oFun = New NContFunciones
    If Not oFun.PermiteModificarAsiento(Format(ldFecha, gsFormatoMovFecha), False) Then
        MsgBox "No se podrá realizar el Proceso ya que la fecha pertenece a un mes ya cerrado", vbInformation, "Aviso"
        txtFecha.SetFocus
        Set oFun = Nothing
        Exit Sub
    End If
    Set oFun = Nothing
    
    If MsgBox("Se va a realizar el registro del Deterioro del Bien" & Chr(10) & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    lsGlosa = Trim(txtGlosa.Text)
    lsPlantilla1 = "44M40RUAG"
    lsPlantilla2 = "18M9090RUAG"
    ReDim MatSoles(7, 0)
    ReDim MatDolares(7, 0)

    Set oMov = New DMov
    oMov.BeginTrans
    bTransMov = True
    
    'Separamos Soles y Dolares
    For i = 1 To FeAdj.Rows - 1
        If FeAdj.TextMatrix(i, 1) = "." And CCur(Trim(FeAdj.TextMatrix(i, 7))) > 0 Then
            If FeAdj.TextMatrix(i, 14) = "1" Then
                index = UBound(MatSoles, 2) + 1
                ReDim Preserve MatSoles(7, index)
                MatSoles(1, index) = CLng(FeAdj.TextMatrix(i, 10)) 'nMovNroAF
                MatSoles(2, index) = Trim(FeAdj.TextMatrix(i, 2)) 'lsSerie
                MatSoles(3, index) = Trim(FeAdj.TextMatrix(i, 12)) 'lsBSCod
                MatSoles(4, index) = Trim(FeAdj.TextMatrix(i, 9)) 'lsAgeCod
                MatSoles(5, index) = CInt(Trim(FeAdj.TextMatrix(i, 15))) 'lnBANCod
                MatSoles(6, index) = CCur(Trim(FeAdj.TextMatrix(i, 7))) 'lnMontoxDepreciar
                MatSoles(7, index) = CInt(Trim(FeAdj.TextMatrix(i, 16))) 'ValorResidual
            Else
                index = UBound(MatDolares, 2) + 1
                ReDim Preserve MatDolares(7, index)
                MatDolares(1, index) = CLng(FeAdj.TextMatrix(i, 10)) 'nMovNroAF
                MatDolares(2, index) = Trim(FeAdj.TextMatrix(i, 2)) 'lsSerie
                MatDolares(3, index) = Trim(FeAdj.TextMatrix(i, 12)) 'lsBSCod
                MatDolares(4, index) = Trim(FeAdj.TextMatrix(i, 9)) 'lsAgeCod
                MatDolares(5, index) = CInt(Trim(FeAdj.TextMatrix(i, 15))) 'lnBANCod
                MatDolares(6, index) = CCur(Trim(FeAdj.TextMatrix(i, 7))) 'lnMontoxDepreciar
                MatDolares(7, index) = CInt(Trim(FeAdj.TextMatrix(i, 16))) 'ValorResidual
            End If
        End If
    Next
    
    For iMoneda = 1 To 2
        If iMoneda = 1 Then
            MatDatos = MatSoles
        Else
            MatDatos = MatDolares
        End If
        If UBound(MatDatos, 2) > 0 Then
            For iSleep = 0 To Rnd(2000) * 1000
            Next
            lsMovNro = oMov.GeneraMovNro(ldFecha, Right(gsCodAge, 2), gsCodUser)
            oMov.InsertaMov lsMovNro, gnDeterioroAF, Left(lsGlosa, 250), gMovEstContabMovContable, gMovFlagVigente
            lnMovNro = oMov.GetnMovNro(lsMovNro)
            lsMovNroTotal = lsMovNroTotal & lsMovNro & ","
            lnMovItem = 0
            lnNroReg = UBound(MatDatos, 2)

            For index = 1 To lnNroReg
                lnMovNroAF = MatDatos(1, index)
                lsSerie = MatDatos(2, index)
                lsBSCod = MatDatos(3, index)
                lsAgeCod = Format(MatDatos(4, index), "00")
                lnBANCod = MatDatos(5, index)
                lnMontoxDepreciar = MatDatos(6, index)
                lnValorResidual = MatDatos(7, index)
            
                oMov.InsertaMovBSAF Year(ldFecha), lnMovNroAF, index, lsBSCod, lsSerie, lnMovNro
            
                lnMovItem = lnMovItem + 1
                lsCtaCont = ReemplazaPlantilla(lsPlantilla1, lsAgeCod, iMoneda, lnBANCod)
                oMov.InsertaMovCta lnMovNro, lnMovItem, lsCtaCont, (lnMontoxDepreciar - lnValorResidual)
                
                lsCtaCont = ReemplazaPlantilla(lsPlantilla2, lsAgeCod, iMoneda, lnBANCod)
                oMov.InsertaMovCta lnMovNro, lnMovItem + lnNroReg, lsCtaCont, (lnMontoxDepreciar - lnValorResidual) * -1
            Next
            If iMoneda = 2 Then
                oMov.GeneraMovME lnMovNro, lsMovNro
            End If
        End If
    Next
    
    oMov.CommitTrans
    bTransMov = False
        
    Screen.MousePointer = 0
    MsgBox "Se ha registrado el Deterioro de los Bienes con Éxito", vbInformation, "Aviso"
    
    Set oPrevio = New clsPrevio
    Set oAsiento = New NContImprimir
    Movs = Split(lsMovNroTotal, ",")

    For i = 0 To UBound(Movs) - 1
        oPrevio.Show oAsiento.ImprimeAsientoContable(Movs(i), 60, 80, Caption), Caption, True
    Next
    
    Limpiar
    cmdBuscar_Click
    
    Set oMov = Nothing
    Set oAsiento = Nothing
    Set oPrevio = Nothing
    Exit Sub
ErrGrabar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
    If bTransMov Then
        oMov.RollbackTrans
        Set oMov = Nothing
    End If
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Function ReemplazaPlantilla(ByVal psPlantilla As String, ByVal psAgeCod As String, ByVal pnMoneda As Integer, ByVal pnBANCod As Integer) As String
    Dim obj As New DBien
    Dim lsRubro As String
    lsRubro = obj.RecuperaSubCtaxBANCod(pnBANCod)
    ReemplazaPlantilla = psPlantilla
    ReemplazaPlantilla = Replace(ReemplazaPlantilla, "M", pnMoneda)
    ReemplazaPlantilla = Replace(ReemplazaPlantilla, "RU", lsRubro)
    ReemplazaPlantilla = Replace(ReemplazaPlantilla, "AG", psAgeCod)
    Set obj = Nothing
End Function
