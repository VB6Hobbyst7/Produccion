VERSION 5.00
Object = "{160AE063-3670-11D5-8214-000103686C75}#6.0#0"; "PryOcxExplorer.ocx"
Begin VB.Form FrmCredTraspCartera 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferencia de Cartera entre Agencias"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9585
   Icon            =   "FrmCredTraspCartera.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   9255
      Begin SICMACT.FlexEdit FETrasCarte 
         Height          =   4455
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7858
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Nro-Cuenta Antigua-Cliente-Cuenta Nueva"
         EncabezadosAnchos=   "400-2000-4400-2000"
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
         ColumnasAEditar =   "X-X-X-X"
         ListaControles  =   "0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-C-L-C"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "Nro"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame frmAgencias 
      Caption         =   "Agencias"
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
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   9255
      Begin SICMACT.TxtBuscar TxtAgencia 
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
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
         sTitulo         =   ""
      End
      Begin VB.Label lblAgencia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   6645
      End
   End
   Begin VB.Frame frmTipoTraspaso 
      Caption         =   "Tipo de transferencia de cartera"
      Enabled         =   0   'False
      Height          =   615
      Left            =   240
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   8895
      Begin VB.ComboBox cboProductos 
         Height          =   315
         Left            =   5160
         TabIndex        =   14
         Top             =   240
         Width           =   3495
      End
      Begin VB.CheckBox chkTipoTrans 
         Caption         =   "Activar reclasificación de creditos"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Producto===>"
         Height          =   255
         Left            =   3960
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   11
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Salir"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   5
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CommandButton cmdGeneCod 
      Caption         =   "Transferir Créditos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   4
      Top             =   7320
      Width           =   1815
   End
   Begin PryOcxExplorer.OcxCdlgExplorer CdlgFile 
      Height          =   615
      Left            =   7800
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      Filtro          =   ""
      Altura          =   0
   End
   Begin VB.CommandButton cmdCargarArch 
      Caption         =   "Cargar Archivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Frame frmAnalista 
      Caption         =   "Cargar Analista"
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
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      Begin VB.ComboBox cboAnalistas 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   8535
      End
      Begin SICMACT.TxtBuscar txtBuscaPersona 
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
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
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.Label lblAnalista 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   6645
      End
   End
End
Attribute VB_Name = "FrmCredTraspCartera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sMatCtasAnti() As String
Dim nPost As Integer
Dim nIndtempo As Integer
'Variables de Excel
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim lsArchivo As String
Dim lsHoja         As String
Dim lbLibroOpen As Boolean
'ALPA 20081212**************************
Dim nTipoTran As Integer
Dim sAgencia As String
'**************************************

'ALPA 20080924********************************************************
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_EXPLORER = &H80000

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Dim ofn As OPENFILENAME
'*********************************************************************

Private Sub chkTipoTrans_Click()
    If chkTipoTrans.value = 0 Then
        nTipoTran = 1
        cboProductos.Enabled = False
    Else
        nTipoTran = 2
        cboProductos.Enabled = True
    End If
End Sub

Private Sub cmdCancelar_Click()
    nPost = 0
    Unload Me
End Sub

Private Sub cmdCargarArch_Click()
    
    Dim oComF As COMFunciones.FCOMExcel
    Dim sRuta As String
    Dim psArchivoALeer As String
    Dim psArchivoAGrabar As String
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja As Excel.Worksheet
    Dim bExiste As Boolean
    Dim bEncontrado As Boolean
    Dim sNAntiguo As String
    Dim sNombreCli As String
    Dim sNNuevo As String
    Dim j As Integer
    Dim fs As New Scripting.FileSystemObject
    'ALPA 20080924*********************************************************
    Dim Filename As String
    Filename = OpenFile(Me.hwnd, "Archivo de texto|*.xls", "Abrir documento", vbNullString)
    If Filename = "" Then
        Exit Sub
    Else
        psArchivoALeer = Filename
    End If
    '**********************************************************************
    'On Error GoTo ErrBegin
    'psArchivoALeer = App.path & "\Spooler\traspaso.xls"
    'psArchivoALeer = "C:\traspaso.xls"
    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoALeer)
    Set oComF = New COMFunciones.FCOMExcel
    bEncontrado = False
    For Each xlHoja In xlLibro.Worksheets
        If UCase(xlHoja.Name) = UCase("Traspaso") Then
            bEncontrado = True
            xlHoja.Activate
            Exit For
        End If
    Next
    If bEncontrado = False Then
        oComF.ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlHoja, True
        MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
        Exit Sub
    End If
        'FETrasCarte.Clear
        Dim i As Integer
        If nPost > 0 Then
            For i = 1 To nPost
                FETrasCarte.EliminaFila (1)
            Next i
        End If
        nPost = 0
        For j = 1 To 10000
            sNAntiguo = Format(xlHoja.Range("A" & (j + 1)), "000000000000000000")
            sNombreCli = xlHoja.Range("B" & (j + 1))
            sNNuevo = Format(xlHoja.Range("C" & (j + 1)), "000000000000000000")
            If sNAntiguo = "" Then
              Exit For
            End If
            FETrasCarte.AdicionaFila
            If sNNuevo = "" Then
            sNNuevo = ""
            End If
            FETrasCarte.TextMatrix(j, 0) = j
            FETrasCarte.TextMatrix(j, 1) = sNAntiguo
            FETrasCarte.TextMatrix(j, 2) = sNombreCli
            FETrasCarte.TextMatrix(j, 3) = ""
            ReDim Preserve sMatCtasAnti(1 To 3, 1 To j)
            nPost = j
            sMatCtasAnti(1, j) = sNAntiguo
            sMatCtasAnti(2, j) = sNombreCli
            sMatCtasAnti(3, j) = ""
        Next j
        cmdGeneCod.Enabled = True
        cmdCargarArch.Enabled = False
'ErrBegin:
'    oComF.ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlHoja, True
End Sub

Private Sub cmdGeneCod_Click()

    Dim odCCred As COMDCredito.DCOMCredito
    Dim lsPersNombre As String
    Dim j As Integer
    If nTipoTran = 1 Then
           sAgencia = TxtAgencia.Text
           If cboAnalistas.Text = "" Then
                MsgBox "Seleccione Analista ", vbCritical, "Error"
                Exit Sub
           End If
           If TxtAgencia.Text = "" Then
                MsgBox "Seleccione Agencia ", vbCritical, "Error"
                Exit Sub
           End If
    Else
        
        If cboProductos.Text = "" Then
            MsgBox "Seleccione Productos ", vbCritical, "Error"
            Exit Sub
        End If
    End If
    
    Dim psMovAct As String
    Dim loContFunct As COMNContabilidad.NCOMContFunciones
   
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
    psMovAct = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, "SIST") ' gsCodUser)
    Set loContFunct = Nothing
    
    Set odCCred = New COMDCredito.DCOMCredito
If nPost >= 1 Then
    For j = 1 To nPost
    lsPersNombre = ""
    sAgencia = ""
    If nTipoTran = 2 Then
        sAgencia = Mid(sMatCtasAnti(1, j), 4, 2)
    Else
         sAgencia = Right(TxtAgencia.Text, 2)
    End If
        sMatCtasAnti(3, j) = odCCred.TransferenciaCar(sMatCtasAnti(1, j), Right(cboAnalistas.Text, 13), gdFecSis, gsCodUser, gsCodCMAC, sAgencia, Mid(sMatCtasAnti(1, j), 6, 3), Mid(sMatCtasAnti(1, j), 9, 1), psMovAct, nTipoTran, IIf(Right(cboProductos.Text, 1) = "", 0, Right(cboProductos.Text, 1)), lsPersNombre)
        sMatCtasAnti(2, j) = lsPersNombre
    Next j
    'FETrasCarte.Clear**
    Dim i As Integer
    If nPost > 0 Then
        For i = 1 To nPost
            FETrasCarte.EliminaFila (1)
        Next i
    End If
    
    For j = 1 To nPost
        FETrasCarte.AdicionaFila
        FETrasCarte.TextMatrix(j, 0) = j
        FETrasCarte.TextMatrix(j, 1) = sMatCtasAnti(1, j)
        FETrasCarte.TextMatrix(j, 2) = sMatCtasAnti(2, j)
        FETrasCarte.TextMatrix(j, 3) = sMatCtasAnti(3, j)
    Next j
Else
    MsgBox "No se encuentra ninguna cuenta que transferir", vbCritical, "Transferencia de cartera"
    Exit Sub
End If

End Sub



Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Dim oGen As COMDConstSistema.DCOMGeneral
    Dim oCons As COMDConstantes.DCOMConstantes
    Set oCons = New COMDConstantes.DCOMConstantes
    Dim oconecta As COMConecta.DCOMConecta
    Dim oDCred As COMDCredito.DCOMCredito
    Dim rsAnalista As ADODB.Recordset
    Dim rsTipoProd As ADODB.Recordset
    Dim sAnalistas As String
    Dim ssql As String
    Dim bMuestraSoloAnalistaActual As Integer
    Set oconecta = New COMConecta.DCOMConecta
    oconecta.AbreConexion
    cmdGeneCod.Enabled = False
    Set oGen = New COMDConstSistema.DCOMGeneral
    sAnalistas = oGen.LeeConstSistema(gConstSistRHCargoCodAnalistas)
    bMuestraSoloAnalistaActual = oGen.LeeConstSistema(58)
    Set oGen = Nothing
    
    ssql = "Select R.cPersCod, P.cPersNombre from RRHH R inner join Persona P ON R.cPersCod = P.cpersCod "
    ssql = ssql & " AND nRHEstado in (201,301) "
    ssql = ssql & " inner join RHCargos RC ON R.cPersCod = RC.cPersCod "
    ssql = ssql & " where  RC.cRHCargoCod in (" & sAnalistas & ") AND RC.dRHCargoFecha = (select MAX(dRHCargoFecha) from RHCargos RHC2 where RHC2.cPersCod = RC.cPersCod) "
    'sSQL = sSQL & " and R.cAgenciaActual='" & psCodAge & "'"
    ssql = ssql & " order by P.cPersNombre "
        
    Set rsAnalista = oconecta.CargaRecordSet(ssql)
    oconecta.CierraConexion
    Set oconecta = Nothing
    '****
    nIndtempo = -99
    Do While Not rsAnalista.EOF
        If bMuestraSoloAnalistaActual = 1 Then
            If rsAnalista!cPersCod = gsCodPersUser Then
                nIndtempo = rsAnalista.AbsolutePosition - 1
            End If
        End If
        cboAnalistas.AddItem PstaNombre(rsAnalista!cPersNombre) & Space(200) & rsAnalista!cPersCod
        rsAnalista.MoveNext
    Loop
    
    If bMuestraSoloAnalistaActual = 1 Then
        cboAnalistas.Enabled = False
        cboAnalistas.ListIndex = nIndtempo
    End If
   '*****

    Me.TxtAgencia.rs = oCons.getAgencias(, , True)
    
    'ALPA 20081212******************************************
    Set oDCred = New COMDCredito.DCOMCredito
    Set rsTipoProd = oDCred.RecuperaTipoProducto
    '
    Do While Not rsTipoProd.EOF
        cboProductos.AddItem rsTipoProd!cConsDescripcion & Space(200) & rsTipoProd!cTipoProd
        rsTipoProd.MoveNext
    Loop
    '
    Set oDCred = Nothing
    nTipoTran = 1
    '*******************************************************
End Sub

Private Sub TxtAgencia_EmiteDatos()
'oGen = New COMDConstSistema.DCOMGeneral
'Set oGen = New COMDConstSistema.DCOMGeneral
Me.lblAgencia.Caption = TxtAgencia.psDescripcion
End Sub
Private Sub txtBuscaPersona_EmiteDatos()
    'Dim oPersTemp As UPersona_Cli
    lblAnalista.Caption = Trim(txtBuscaPersona.psDescripcion)
End Sub
Private Sub cmdImprimir_Click()
    Dim lscCtaCod As String
    Dim nCont As Integer
    Dim nCtaCodNueva As String
    nCont = 0
    lscCtaCod = ""
    Dim i As Integer
    i = 0
    If nPost >= 1 Then
        For i = 1 To nPost
            If sMatCtasAnti(3, i) <> "" Then
            If sMatCtasAnti(3, i) = "No tiene equivalencia" Then
                nCtaCodNueva = "0"
            Else
                nCtaCodNueva = sMatCtasAnti(3, i)
            End If
                If i = 1 Then
                    lscCtaCod = lscCtaCod & nCtaCodNueva
                Else
                    lscCtaCod = lscCtaCod & "," & nCtaCodNueva
                End If
                nCont = nCont + 1
            End If
            nCtaCodNueva = ""
        Next i
        If nCont > 0 Then
            Call GenRepMatrixTransferenciaCartera(lscCtaCod)
        Else
            MsgBox "Lista no se encuentra Transferida", vbCritical, "Aviso!!!"
        End If
    Else
        MsgBox "Debe Cargar el archivo", vbCritical, "Aviso!!!"
    End If
End Sub
Public Sub Reporte_GenRepMatrixTransferenciaCartera(ByVal lsCtaCod As Variant)
    Dim i As Integer
    Dim matTem() As String
    Dim oNcCredito As COMNCredito.NCOMCredito
    Set oNcCredito = New COMNCredito.NCOMCredito
    matTem = oNcCredito.ReportesTransferenciaCartera(matTem, lsCtaCod)
    
    xlHoja1.Range("A1:S100").EntireColumn.Font.FontStyle = "Arial"
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 75
    xlHoja1.PageSetup.TopMargin = 2
    xlHoja1.Range("A9:Z1").EntireColumn.Font.Size = 7
    xlHoja1.Range("A1:A1").RowHeight = 17
    xlHoja1.Range("A1:A1").ColumnWidth = 12
    xlHoja1.Range("B1:B1").ColumnWidth = 12
    xlHoja1.Range("C1:C1").ColumnWidth = 40
    xlHoja1.Range("A1:K1").Merge
    xlHoja1.Range("A1:K1").HorizontalAlignment = xlCenter
    xlHoja1.Cells(1, 1) = "REPORTE DE CUENTAS TRANSFERIDAS"
    xlHoja1.Range("A1:K1").Font.Bold = True
    xlHoja1.Range("A2:A2").Font.Bold = True
    xlHoja1.Range("A3:A3").Font.Bold = True
    xlHoja1.Cells(2, 1) = "FECHA  :"
    xlHoja1.Cells(2, 2) = Format(gdFecSis, "yyyy/mm/dd")
    xlHoja1.Cells(3, 1) = "USUARIO:"
    xlHoja1.Cells(3, 2) = gsCodUser
    xlHoja1.Cells(5, 1) = "NUEVA CUENTA"
    xlHoja1.Cells(5, 2) = "CUENTA ANTIGUA"
    xlHoja1.Cells(5, 3) = "CLIENTE"
    xlHoja1.Cells(5, 4) = "PRODUCTO"
    xlHoja1.Cells(5, 5) = "ESTADO"
    xlHoja1.Cells(5, 6) = "MONEDA"
    xlHoja1.Cells(5, 7) = "LINEA"
    xlHoja1.Cells(5, 8) = "PLAZO"
    xlHoja1.Cells(5, 9) = "SALDO"
    xlHoja1.Cells(5, 10) = "MONTOCOL"
    xlHoja1.Cells(5, 11) = "IMPORTE"
    xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 11)).Interior.Color = RGB(166, 166, 166)
    xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 11)).Borders.LineStyle = 1
    xlHoja1.Range("A5:K5").Font.Bold = True
    If matTem(0, 1) > 0 Then
        For i = 1 To matTem(0, 1)
            xlHoja1.Range(xlHoja1.Cells(i + 5, 1), xlHoja1.Cells(i + 5, 11)).Borders.LineStyle = 1
            xlHoja1.Cells(i + 5, 1) = matTem(1, i)
            xlHoja1.Cells(i + 5, 2) = matTem(2, i)
            xlHoja1.Cells(i + 5, 3) = matTem(3, i)
            xlHoja1.Cells(i + 5, 4) = matTem(4, i)
            xlHoja1.Cells(i + 5, 5) = matTem(5, i)
            xlHoja1.Cells(i + 5, 6) = matTem(6, i)
            xlHoja1.Cells(i + 5, 7) = matTem(7, i)
            xlHoja1.Cells(i + 5, 8) = matTem(8, i)
            xlHoja1.Cells(i + 5, 9) = matTem(9, i)
            xlHoja1.Cells(i + 5, 10) = matTem(10, i)
            xlHoja1.Cells(i + 5, 11) = matTem(11, i)
            'xlHoja1.Cells(xlHoja1.Cells(i + 5, 9), xlHoja1.Cells(i + 5, 11)).NumberFormat = "#,##0.00;-#,##0.00"
        Next i
    End If
    With xlHoja1.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
    
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 55
    End With
End Sub
Sub GenRepMatrixTransferenciaCartera(ByVal lsCtaCod As Variant)
    
        If nPost >= 1 Then
        '**ALPA**20080412**
            Dim sTxtFec As Date
            lsArchivo = App.Path & "\SPOOLER\Transferencia_" & Format(gdFecSis, "yyyymmdd") & Format(Time(), "HHMMSS") & ".XLS"
            lbLibroOpen = gFunGeneral.ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
            lsHoja = "HoTransfCartera"
            ExcelAddHoja lsHoja, xlLibro, xlHoja1
            xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1000, 3)).NumberFormat = "@"
            Call Reporte_GenRepMatrixTransferenciaCartera(lsCtaCod)
            gFunGeneral.ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
            gFunGeneral.CargaArchivo lsArchivo, App.Path & "\SPOOLER\"
        End If
End Sub

'Muestra el cuadro de dialogo para abrir archivos:
Public Function OpenFile(hwnd As Long, Filter As String, Title As String, InitDir As String, Optional Filename As String, Optional FilterIndex As Long) As String
    On Local Error Resume Next

    Dim ofn As OPENFILENAME
    Dim a As Long
    
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = hwnd
    ofn.hInstance = App.hInstance
    
    If VBA.Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
    
    For a = 1 To Len(Filter)
        If Mid$(Filter, a, 1) = "|" Then Mid(Filter, a, 1) = Chr(0)
    Next
    
        ofn.lpstrFilter = Filter
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = InitDir
        If Not Filename = vbNullString Then ofn.lpstrFile = Filename & Space$(254 - Len(Filename))
        ofn.nFilterIndex = FilterIndex
        ofn.lpstrTitle = Title
        ofn.Flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
        a = GetOpenFileName(ofn)

        If a Then
             OpenFile = Trim$(ofn.lpstrFile)
             If VBA.Right$(VBA.Trim$(OpenFile), 1) = Chr(0) Then OpenFile = VBA.Left$(VBA.Trim$(ofn.lpstrFile), Len(VBA.Trim$(ofn.lpstrFile)) - 1)
             
        Else
             OpenFile = vbNullString
             
        End If
        
End Function

'Extrae la extension seleccionada del filtro:
Private Function GetExtension(sfilter As String, Pos As Long) As String
    Dim Ext() As String
    
    Ext = Split(sfilter, vbNullChar)
    
    If Pos = 1 And Ext(Pos) <> "*.*" Then
        GetExtension = "." & Replace(Ext(Pos), "*.", "")
        Exit Function
        
    End If
    
    If Pos = 1 And Ext(Pos) = "*.*" Then
        GetExtension = vbNullString
        Exit Function
        
    End If
    
    If InStr(Ext(Pos + 1), "*.*") Then
       GetExtension = vbNullString
       
    Else
       GetExtension = "." & Replace(Ext(Pos + 1), "*.", "")
       
    End If
    
End Function

