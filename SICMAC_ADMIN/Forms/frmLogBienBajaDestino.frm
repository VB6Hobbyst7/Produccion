VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogBienBajaDestino 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Activo Fijo dados de Baja"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12240
   Icon            =   "frmLogBienBajaDestino.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   12240
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Destino del Bien dado de Baja"
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
      TabIndex        =   9
      Top             =   5685
      Width           =   12135
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   6600
         MaxLength       =   450
         TabIndex        =   11
         Top             =   240
         Width           =   5385
      End
      Begin VB.ComboBox cboDestino 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   3135
      End
      Begin MSComCtl2.DTPicker txtFecha 
         Height          =   315
         Left            =   3960
         TabIndex        =   12
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
      Begin VB.Label lblComentario 
         Caption         =   "Descripción:"
         Height          =   210
         Left            =   5640
         TabIndex        =   14
         Top             =   285
         Width           =   900
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha:"
         Height          =   225
         Left            =   3360
         TabIndex        =   13
         Top             =   285
         Width           =   570
      End
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
      Left            =   9880
      TabIndex        =   4
      Top             =   6465
      Width           =   1125
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
      TabIndex        =   5
      Top             =   6465
      Width           =   1125
   End
   Begin VB.Frame fraOpe 
      Caption         =   "Baja de Activo Fijo"
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
      TabIndex        =   6
      Top             =   60
      Width           =   12120
      Begin VB.TextBox txtSerieNombre 
         Appearance      =   0  'Flat
         Height          =   320
         Left            =   3690
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   960
         Width           =   4860
      End
      Begin VB.TextBox txtBSNombre 
         Appearance      =   0  'Flat
         Height          =   320
         Left            =   3690
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   600
         Width           =   4860
      End
      Begin VB.TextBox txtAreaAgeNombre 
         Appearance      =   0  'Flat
         Height          =   320
         Left            =   3690
         Locked          =   -1  'True
         TabIndex        =   16
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
         TabIndex        =   2
         Top             =   940
         Width           =   1050
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
         TabIndex        =   18
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
         TabIndex        =   19
         Top             =   990
         Width           =   450
      End
      Begin VB.Label lblBien 
         AutoSize        =   -1  'True
         Caption         =   "Categoría :"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   630
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Área/Agencia :"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   270
         Width           =   1080
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
      Left            =   8705
      TabIndex        =   3
      Top             =   6465
      Width           =   1125
   End
   Begin Sicmact.FlexEdit FeAdj 
      Height          =   4095
      Left            =   60
      TabIndex        =   15
      Top             =   1560
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   7223
      Cols0           =   15
      HighLight       =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      EncabezadosNombres=   "Nº-OK-Serie-Descripcion-Fecha-Valor Ini.-Valor Depre.-Por Depre.-Area-Agencia-nMovNro-nAnio-cBSCod-nMovNroAC-cBaja"
      EncabezadosAnchos=   "400-400-1800-3500-1200-1200-1200-1200-800-800-0-0-1200-0-0"
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
      ColumnasAEditar =   "X-1-X-X-X-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-4-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-L-R-R-R-R-C-C-C-C-C-R-L"
      FormatosEdit    =   "0-0-0-0-5-2-2-2-0-0-0-0-0-3-0"
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
Attribute VB_Name = "frmLogBienBajaDestino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************
'** Nombre : frmLogBienBajaDestino
'** Descripción : Destino de los Bienes dados de Baja creado segun ERS059-2013
'** Creación : EJVG, 20130627 09:00:00 AM
'*****************************************************************************
Option Explicit

Private Sub Form_Load()
    CentraForm Me
    CargaCombos
    CargaTxt
    txtFecha.value = Format(gdFecSis, "dd/mm/yyyy")
End Sub
Private Sub CargaTxt()
    Dim oBien As New DBien
    Dim oArea As New DActualizaDatosArea
    txtAreaAgeCod.rs = oArea.GetAgenciasAreas()
    txtBSCod.rs = oBien.GetAFBienesBaja()
    txtSerieCod.rs = oBien.RecuperaSeriesBajadosPaObjeto()
    Set oBien = Nothing
    Set oArea = Nothing
End Sub
Private Sub CargaCombos()
    Dim oConst As New DConstante
    Dim rsDestino As New ADODB.Recordset
    Set rsDestino = oConst.RecuperaConstantes(10020)
    CargaCombo rsDestino, cboDestino, 150
End Sub
Private Sub txtAreaAgeCod_EmiteDatos()
    Dim oBien As New DBien
    Dim lsAreaAgeCod As String
    
    Screen.MousePointer = 11
    txtAreaAgeNombre.Text = ""
    If txtAreaAgeCod.Text <> "" Then
        txtAreaAgeNombre.Text = txtAreaAgeCod.psDescripcion
        lsAreaAgeCod = Left(txtAreaAgeCod.Text, 3) & IIf(Mid(txtAreaAgeCod.Text, 4, 2) = "", "01", Mid(txtAreaAgeCod.Text, 4, 2))
    End If
    txtBSCod.Text = ""
    txtBSNombre.Text = ""
    txtBSCod.rs = oBien.GetAFBienesBaja(lsAreaAgeCod)
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
    txtSerieCod.rs = oBien.RecuperaSeriesBajadosPaObjeto(lsAreaAgeCod, txtBSCod.Text, True)
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
    If txtAreaAgeCod.Text <> "" Then
        lsAreaAgeCod = Left(txtAreaAgeCod.Text, 3) & IIf(Mid(txtAreaAgeCod.Text, 4, 2) = "", "01", Mid(txtAreaAgeCod.Text, 4, 2))
    End If
    Set rsSeries = oBien.GetAFBienesBajadosPaDestino(lsAreaAgeCod, txtBSCod.Text, txtSerieCod.Text)
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
Private Sub cmdGrabar_Click()
    Dim oBien As DBien
    Dim i As Long
    Dim lbOK As Boolean
    Dim lsValFecha As String
    Dim lnMovNro As Long
    Dim lnDestino As Integer
    Dim ldBaja As Date, ldFecha As Date
    Dim lsDescripcion As String
    Dim bTrans As Boolean

    On Error GoTo ErrGrabar
    
    If FlexVacio(FeAdj) Then
        MsgBox "Ud. debe de seleccionar los registros que se van a registrar su destino de Baja", vbInformation, "Aviso"
        Exit Sub
    Else
        For i = 1 To FeAdj.Rows - 1
            If FeAdj.TextMatrix(i, 1) = "." Then
                lbOK = True
                Exit For
            End If
        Next
        If Not lbOK Then
            MsgBox "Ud. debe de seleccionar los registros que se van a registrar su destino de Baja", vbInformation, "Aviso"
            Exit Sub
        End If
        If cboDestino.ListIndex = -1 Then
            MsgBox "Ud. debe de seleccionar el destino para los Bienes", vbInformation, "Aviso"
            cboDestino.SetFocus
            Exit Sub
        End If
        lsValFecha = ValidaFecha(txtFecha.value)
        If lsValFecha <> "" Then
            MsgBox lsValFecha, vbInformation, "Aviso"
            txtFecha.SetFocus
            Exit Sub
        End If
        If Len(Trim(txtDescripcion.Text)) = 0 Then
            MsgBox "Ud. de ingresar la Descripcion del Destino de la Baja", vbInformation, "Aviso"
            txtDescripcion.SetFocus
            Exit Sub
        End If
    End If
    
    If MsgBox("Se va a realizar el registro del Destino de la Baja del Bien" & Chr(10) & "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then Exit Sub
    
    lnDestino = CInt(Trim(Right(cboDestino.Text, 3)))
    ldFecha = CDate(txtFecha.value)
    lsDescripcion = Trim(txtDescripcion.Text)
    
    Set oBien = New DBien
    oBien.dBeginTrans
    bTrans = True
    
    For i = 1 To FeAdj.Rows - 1
        If FeAdj.TextMatrix(i, 1) = "." Then
            lnMovNro = CLng(FeAdj.TextMatrix(i, 10))
            ldBaja = CDate(FeAdj.TextMatrix(i, 13))
            If ldFecha < ldBaja Then
                MsgBox "La Fecha es menor que la Fecha que se dió de Baja a la Serie " & FeAdj.TextMatrix(i, 2), vbCritical, "Aviso"
                oBien.dRollbackTrans
                Set oBien = Nothing
                Exit Sub
            End If
            'oBien.InsertaDestinoBajaAF lnMovNro, lnDestino, ldFecha, lsDescripcion
            oBien.ActualizaDestinoBajaAF lnMovNro, lnDestino, ldFecha, lsDescripcion
        End If
    Next
    
    oBien.dCommitTrans
    bTrans = False
    
    MsgBox "Se ha registrado el Destino de los Bienes con Éxito", vbInformation, "Aviso"
    cboDestino.ListIndex = -1
    txtFecha.value = Format(gdFecSis, "dd/mm/yyyy")
    txtDescripcion.Text = ""
    cmdBuscar_Click
    Exit Sub
ErrGrabar:
    MsgBox Err.Description, vbCritical, "Aviso"
    If bTrans Then
        oBien.dRollbackTrans
        Set oBien = Nothing
    End If
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
    
    lsArchivo = "\spooler\RptDestinoBajaActivos" & UCase(gsCodUser) & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    Set xlsLibro = xlsAplicacion.Workbooks.Add

    Set xlsHoja = xlsLibro.Worksheets.Add
    xlsHoja.Name = "Reporte Destino Baja de Activos"
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
Private Sub cmdSalir_Click()
    Unload Me
End Sub
