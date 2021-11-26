VERSION 5.00
Begin VB.Form frmAdmCheckListDocument 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CheckList Documentario"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9960
   Icon            =   "frmAdmCheckListDocument.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frNiveles 
      Caption         =   "Niveles"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   9735
      Begin VB.ComboBox cmbNiveles 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   9495
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   300
      Left            =   7440
      TabIndex        =   6
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   300
      Left            =   8640
      TabIndex        =   5
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenerarExcel 
      Caption         =   "Generar PDF"
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   9735
      Begin SICMACT.FlexEdit feRequisitos 
         Height          =   2985
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   5265
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Requisito-Estado-idRequisito"
         EncabezadosAnchos=   "300-8000-800-0"
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
         ColumnasAEditar =   "X-X-2-X"
         ListaControles  =   "0-0-3-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-C"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Categoría "
      ForeColor       =   &H00FF0000&
      Height          =   680
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin VB.ComboBox cboCategoria 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   9495
      End
   End
End
Attribute VB_Name = "frmAdmCheckListDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmAdmCheckLsitDocument
'** Descripción : Formulario que permite registrar los requisitos solicitados por crédito
'** Creación    : RECO, 20150421 - ERS010-2015
'**********************************************************************************************
Option Explicit
Dim sCtaCod As String
Dim sTpoCred As String
Dim vMatriz() As Variant
Dim nInicializa As Integer
Dim nTpoOpe As Integer
Dim bCheck As Boolean

Dim sTpoProd As String 'JOEP20190102
Dim nMonto As Currency  'JOEP20190102
Dim sTpoCateg As String 'JOEP20190102
Dim bCargIni As Boolean

Public Enum TipoOperacionCheckList
    nRegCntrlPre = 1
    nRegCntrlPost = 2
    nCheckListMant = 3
    nCheckListConsul = 4
    nRegSugerencia = 5
    nRegAprobacion = 6
    nRegSugApro = 7
    nRegSugerenciaCF = 8
    nRegRenovacionCF = 9
End Enum

'Public Function Inicio(ByVal psCtaCod As String, ByVal psTpoCred As String, ByVal pnTpoOpe As TipoOperacionCheckList) As Boolean'Comento JOEP20181229 CP
Public Function Inicio(ByVal psCtaCod As String, ByVal psTpoCateg As String, ByVal psTpoProd As String, ByVal pnMonto As Currency, ByVal psTpoCred As String, ByVal pnTpoOpe As TipoOperacionCheckList) As Boolean

    nInicializa = 1
    'JOEP20181222 CP
    Frame2.Top = 840
    frmAdmCheckListDocument.Height = 5070
    cmdGrabar.Top = 4320
    cmdSalir.Top = 4320
    cmdGenerarExcel.Top = 4320
    
    bCargIni = False
    sTpoCateg = psTpoCateg
    sTpoProd = psTpoProd
    nMonto = pnMonto
    sTpoCred = psTpoCred
    'JOEP20181222 CP
    cmdGenerarExcel.Enabled = False
    bCheck = False
    'Call CargarCombo(psTpoCred) 'comento JOEP20181222 CP
    Call CargarCombo(psCtaCod, psTpoCateg, psTpoProd, pnMonto, psTpoCred, pnTpoOpe)
    nInicializa = 2
    sCtaCod = psCtaCod
    'sTpoCred = psTpoCred 'comento JOEP20181222 CP
    nTpoOpe = pnTpoOpe
    If CargarMatriz = 1 Then
        Call CargaDatoMatriz(cboCategoria.ItemData(cboCategoria.ListIndex))
        Call cboCategoria_Click
        bCargIni = True
        
        'JOEP20190116 CP
        Call LimpiaFlex(feRequisitos)
        If pnTpoOpe = 1 Or pnTpoOpe = 2 Then
            feRequisitos.Enabled = False
            cmdGrabar.Enabled = False
            cmdGenerarExcel.Enabled = True
        End If
        'JOEP20190116 CP
        Me.Show 1
        Inicio = bCheck
    Else
        MsgBox "No se realizó la configuración del CheckList para este Producto.", vbInformation, "Alerta"
     End If
End Function

'Private Sub CargarCombo(ByVal psTpoCred As String)'Comento JOEP20181229 CP
Private Sub CargarCombo(ByVal psCtaCod As String, ByVal psTpoCateg As String, ByVal psTpoProd As String, ByVal pnMonto As Currency, ByVal psTpoCred As String, Optional ByVal psTpoOpe As Integer = 0)
    Dim obj As New COMNCredito.NCOMCredito
    Dim rs As ADODB.Recordset
    Dim i As Integer
    
    'Set rs = obj.ListaCredAdmSecciones(Mid(psTpoCred, 1, 2) & "0")'Comento JOEP20191229 CP
    Set rs = obj.ListaCredAdmSecciones(psCtaCod, psTpoCateg, psTpoProd, pnMonto, psTpoCred, psTpoOpe) 'JOEP20191229 CP
    If Not (rs.BOF And rs.EOF) Then
        For i = 1 To rs.RecordCount
            cboCategoria.AddItem rs!cDescripcion & Space(500) & rs!cItemCantConf
            cboCategoria.ItemData(cboCategoria.NewIndex) = "" & rs!nIdSeccion
            rs.MoveNext
        Next
        cboCategoria.ListIndex = 0
    End If
Set obj = Nothing
RSClose rs
End Sub
        
'Private Sub CargarRequisitos()'Comento JOEP20181229 CP
Private Sub CargarRequisitos(ByVal psTpoCateg As String, ByVal psTpoProd As String, ByVal psItem As String, ByVal pnCantConf As Integer, Optional ByVal nComboSel As Integer = 0)
    Dim obj As New COMNCredito.NCOMCredito
    Dim OCon As New COMDConstantes.DCOMConstantes
    Dim rs As ADODB.Recordset
    Dim i As Integer
    
    feRequisitos.Clear
    FormateaFlex feRequisitos

    If nInicializa = 1 Then
        'Set rs = obj.ListaCredAdmRequisitos(cboCategoria.ItemData(cboCategoria.ListIndex))
        Set rs = obj.ListaCredAdmRequisitos(psTpoCateg, psTpoProd, psItem, pnCantConf)
        feRequisitos.CargaCombo OCon.RecuperaConstantes(10064)
        If Not (rs.EOF And rs.BOF) Then
            For i = 1 To rs.RecordCount
                If InStr(rs!cItem, ".") = 0 Then 'JOEP
                    feRequisitos.AdicionaFila
                    feRequisitos.TextMatrix(i, 3) = rs!nIdRequisito
                    feRequisitos.TextMatrix(i, 2) = "N/A" & Space(75) & "3"
                    feRequisitos.TextMatrix(i, 1) = rs!cDescripcion
                End If 'JOEP
                rs.MoveNext
            Next
        End If
    Else
        'JOEP
        If nComboSel = 1 Then 'Combo Categoria
            Call CargaDatoMatriz(cboCategoria.ItemData(cboCategoria.ListIndex))
        Else 'Combo Niveles
            Call CargaDatoMatriz(cmbNiveles.ItemData(cmbNiveles.ListIndex))
        End If
        'JOEP
        'Call CargaDatoMatriz(cboCategoria.ItemData(cboCategoria.ListIndex))
        
    End If
        
    feRequisitos.row = 1
    feRequisitos.TopRow = 1
End Sub

Private Sub cboCategoria_Click()
Dim obj As COMDCredito.DCOMCredito
Dim sItem As String
Dim nCantConf As Integer
Dim nDato As String
Dim rs As ADODB.Recordset
Dim i As Integer
Set obj = New COMDCredito.DCOMCredito

If cboCategoria.ListIndex = -1 Then Exit Sub

nDato = Trim(Right(cboCategoria.Text, 9))
sItem = Mid(nDato, 1, (InStr(nDato, "-")) - 1)
nCantConf = Mid(nDato, (InStr(nDato, "-")) + 1, 6)

    If nInicializa = 2 Then
        Call ActualizaMatriz
    End If
    
    'Call CargarRequisitos
    Call CargarRequisitos(sTpoCateg, sTpoProd, sItem, nCantConf, 1) 'JOEP20181229 CP

'JOEP20181229 CP

    Set obj = New COMDCredito.DCOMCredito
    Set rs = obj.ListaCredAdmNiveles(sTpoCateg, sTpoProd, sItem, nCantConf)
    
    If Not (rs.BOF And rs.EOF) Then
      
        Frame2.Top = 1440
        frmAdmCheckListDocument.Height = 5700
        cmdGrabar.Top = 4920
        cmdSalir.Top = 4920
        cmdGenerarExcel.Top = 4920
    
        frNiveles.Visible = True
        cmbNiveles.Clear
        For i = 1 To rs.RecordCount
            cmbNiveles.AddItem rs!cDescripcion & Space(500) & rs!cItemCantConf
            cmbNiveles.ItemData(cmbNiveles.NewIndex) = "" & rs!nIdSeccion
            rs.MoveNext
        Next
    cmbNiveles.ListIndex = 0
    Else
        cmbNiveles.Clear
        frNiveles.Visible = False
        Frame2.Top = 840
        frmAdmCheckListDocument.Height = 5070
        cmdGrabar.Top = 4320
        cmdSalir.Top = 4320
        cmdGenerarExcel.Top = 4320
    End If
    
If bCargIni = True Then
    If (nTpoOpe <> 1 And nTpoOpe <> 2) Then
        Set rs = obj.CP_getMensaje(sCtaCod, sTpoCateg, sTpoProd, sItem, nCantConf)
        If Not (rs.BOF And rs.EOF) Then
            For i = 1 To rs.RecordCount
                MsgBox rs!cDescripcion, vbInformation, "Aviso"
                rs.MoveNext
            Next i
        End If
    End If
End If
'JOEP20181229 CP

End Sub

Private Sub cmdGenerarExcel_Click()
    bCargIni = False 'JOEP20190108 CP
    Call ActualizaMatriz
    Call Imprimir
End Sub

Private Sub cmdGrabar_Click()
    Dim obj As New COMNCredito.NCOMCredito
    If Not CP_GetValMsg() Then Exit Sub 'JOEP20181229 CP
    
    If MsgBox("Esta seguro que desea guardar la información.", vbQuestion + vbYesNo) = vbYes Then
        Dim i As Integer, j As Integer
        Dim idEstado As Integer, idrequisito As Integer
        Screen.MousePointer = 11
        cmdGrabar.Enabled = False
        Call ActualizaMatriz
        Call obj.EliminaCheckListCta(sCtaCod)
        For j = 0 To UBound(vMatriz, 2)
            Call obj.RegistraRequisitoCta(sCtaCod, Trim(Right(vMatriz(2, j), 4)), Trim(Right(vMatriz(3, j), 4)), nTpoOpe)
        Next
        Screen.MousePointer = 0
        cmdGenerarExcel.Enabled = True
        feRequisitos.Enabled = False 'JOEP20190116 CP
        bCheck = True
    Else
        bCheck = False
    End If
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Function CargarMatriz() As Integer
    Dim obj As New COMNCredito.NCOMCredito
    Dim rs As New ADODB.Recordset
    Dim i As Integer, nCantidad As Integer
    
    'Set rs = obj.ObtieneRequisitosTpoCred(Mid(sTpoCred, 1, 2) & "0", sCtaCod)'comento JOEP20181222 CP
    Set rs = obj.ObtieneRequisitosTpoCred(sTpoCateg, sTpoProd, sCtaCod, nMonto, sTpoCred, nTpoOpe) 'JOEP20181222 CP
    
    If Not (rs.BOF And rs.EOF) Then
        nCantidad = rs.RecordCount
        ReDim Preserve vMatriz(3, 0 To nCantidad - 1)
        For i = 0 To nCantidad - 1
            vMatriz(1, i) = rs!nIdSeccion
            vMatriz(2, i) = rs!cRequisitoDesc & Space(50) & rs!nIdRequisito
            vMatriz(3, i) = rs!cEstado & Space(75) & rs!nEstado
            rs.MoveNext
        Next
        CargarMatriz = 1
    Else
        CargarMatriz = 0
    End If
End Function

Private Sub ActualizaMatriz()
    Dim i As Integer, j As Integer
    For i = 1 To feRequisitos.rows - 1
        For j = 0 To UBound(vMatriz, 2)
            If feRequisitos.TextMatrix(i, 3) = Trim(Right(vMatriz(2, j), 4)) Then
                vMatriz(3, j) = feRequisitos.TextMatrix(i, 2)
            End If
        Next
    Next
End Sub

Private Sub CargaDatoMatriz(ByVal pnSeccion As Integer)
    Dim i As Integer, j As Integer
    For j = 0 To UBound(vMatriz, 2)
        If pnSeccion = vMatriz(1, j) Then
            feRequisitos.AdicionaFila
            feRequisitos.TextMatrix(feRequisitos.rows - 1, 1) = Trim(Mid(vMatriz(2, j), 1, Len(vMatriz(2, j)) - 4))
            feRequisitos.TextMatrix(feRequisitos.rows - 1, 2) = vMatriz(3, j)
            feRequisitos.TextMatrix(feRequisitos.rows - 1, 3) = Trim(Right(vMatriz(2, j), 4))
        End If
    Next
End Sub

'JOEP20190107 CP
Private Sub cmbNiveles_Click()
Dim sItemN As String
Dim nCantConfN As Integer
Dim nDatoN As String

nDatoN = Trim(Right(cmbNiveles.Text, 9))
sItemN = Mid(nDatoN, 1, (InStr(nDatoN, "-")) - 1)
nCantConfN = Mid(nDatoN, (InStr(nDatoN, "-")) + 1, 6)

If nInicializa = 2 Then
    Call ActualizaMatriz
End If
    
    Call CargarRequisitos(sTpoCateg, sTpoProd, sItemN, nCantConfN, 2) 'JOEP20181229 CP
    
End Sub
'JOEP20190107 CP

Public Sub Imprimir()
'Comento JOEP20190116 CP
'    Dim fs As New Scripting.FileSystemObject
'    Dim xlsAplicacion As New Excel.Application
'    Dim obj As New COMNCredito.NCOMCredito
'    Dim oCred As New COMDCredito.DCOMCreditos
'    Dim rs As New ADODB.Recordset
'    Dim rsObs As New ADODB.Recordset
'    Dim xlsLibro As Excel.Workbook
'    Dim xlHoja1 As Excel.Worksheet
'    Dim lsArchivo As String, lsFile As String, lsNomHoja As String
'    Dim i As Integer, j As Integer, IniTablas As Integer
'    Dim lbExisteHoja As Boolean
'    Dim rsCred As New ADODB.Recordset
'    Set rs = obj.ObtenerDatosCredChekList(sCtaCod)
'    lsNomHoja = "REQUISITOS-LIST"
'    lsFile = "CHECK_LIST_REQUISITOS"
'
'    lsArchivo = "\spooler\" & "CheckList" & "_" & sCtaCod & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xls" 'pti1 Memorandum Nº 1602-2018-GM-DI_CMACM (1)se cambio el formato de xls a xlsx 20/07/2018
'    If fs.FileExists(App.Path & "\FormatoCarta\" & lsFile & ".xls") Then
'        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.Path & "\FormatoCarta\" & lsFile & ".xls")
'    Else
'        MsgBox "No Existe Plantilla en Carpeta FormatoCarta (" & lsFile & ".xls), Consulte con el Area de TI", vbInformation, "Advertencia" 'pti1 Memorandum Nº 1602-2018-GM-DI_CMACM (1)se cambio el formato de xls a xlsx 20/07/2018
'        Exit Sub
'    End If
'
'    For Each xlHoja1 In xlsLibro.Worksheets
'       If xlHoja1.Name = lsNomHoja Then
'            xlHoja1.Activate
'         lbExisteHoja = True
'        Exit For
'       End If
'    Next
'
'    If lbExisteHoja = False Then
'        Set xlHoja1 = xlsLibro.Worksheets
'        xlHoja1.Name = lsNomHoja
'    End If
'
'    Dim nContador As Integer
'    Dim nFilaCan As Integer
'    Dim nFilaAnt As Integer
'    IniTablas = 13
'    nContador = IniTablas
'
'    If Not (rs.EOF And rs.BOF) Then
'        xlHoja1.Cells(5, 4) = rs!cPersNombre
'        xlHoja1.Cells(5, 6) = rs!cPersCodSbs
'        xlHoja1.Cells(5, 13) = rs!cCtaCod
'        xlHoja1.Cells(7, 4) = rs!cUser
'        xlHoja1.Cells(7, 6) = rs!cTpoPrd
'        xlHoja1.Cells(7, 13) = rs!dVigencia
'        xlHoja1.Cells(9, 4) = rs!cAgeDescripcion
'        xlHoja1.Cells(9, 6) = Format(rs!nMontoCol, gsFormatoNumeroView)
'        xlHoja1.Cells(9, 13) = rs!cModalidad
'    End If
'
'    For i = 0 To cboCategoria.ListCount - 1
'
'        cboCategoria.ListIndex = i
'        'nContador = IIf(nContador = 8, 13, nContador) agregado por pti1 18072018 Memorandum Nº 1602-2018-GM-DI/CMACM
'        If ValidaExisteItem(cboCategoria.ListIndex) = True Then
'            xlHoja1.Cells(nContador, 3) = i + 1 & " " & cboCategoria.Text
'            xlHoja1.Range(xlHoja1.Cells(nContador, 3), xlHoja1.Cells(nContador, 18)).Merge
'            xlHoja1.Range(xlHoja1.Cells(nContador, 3), xlHoja1.Cells(nContador, 18)).Style = "Salida"
'            nContador = nContador + 2
'            nFilaAnt = nContador
'            nFilaCan = 1
'
'        End If
'        '********************Agregado por pti1 18072018 Memorandum Nº 1602-2018-GM-DI/CMACM
'        Dim nContadorColumna As Integer
'        Dim nContadorGrupo As Integer
'        Dim nContadorGrupoEstado As Boolean
'
'        nContadorGrupoEstado = False
'        nContadorGrupo = 0
'        nContadorColumna = 0
'
'        For j = 0 To UBound(vMatriz, 2)
'            If Trim(cboCategoria.ItemData(cboCategoria.ListIndex)) = Trim(vMatriz(1, j)) Then
'
'             If Right(vMatriz(3, j), 1) <> 3 Then
'               nContadorColumna = nContadorColumna + 1
'               nContadorGrupo = nContadorGrupo + 1
'               If nContadorColumna = 1 Then
'                        xlHoja1.Cells(nContador, 2) = nContadorGrupo
'                        xlHoja1.Cells(nContador, 3) = Trim(Mid(vMatriz(2, j), 1, Len(vMatriz(2, j)) - 4))
'                        xlHoja1.Range(xlHoja1.Cells(nContador, 3), xlHoja1.Cells(nContador, 3 + 2)).Merge
'                        xlHoja1.Cells(nContador, 6) = Trim(Mid(vMatriz(3, j), 1, Len(vMatriz(3, j)) - 1))
'                        xlHoja1.Range(xlHoja1.Cells(nContador, 6), xlHoja1.Cells(nContador, 6)).Borders.LineStyle = 1
'                        nContadorGrupoEstado = True
'               End If
'               If nContadorColumna = 2 Then
'                        xlHoja1.Cells(nContador, 8) = nContadorGrupo
'                        xlHoja1.Cells(nContador, 9) = Trim(Mid(vMatriz(2, j), 1, Len(vMatriz(2, j)) - 4))
'                        xlHoja1.Range(xlHoja1.Cells(nContador, 9), xlHoja1.Cells(nContador, 9 + 2)).Merge
'                        xlHoja1.Cells(nContador, 12) = Trim(Mid(vMatriz(3, j), 1, Len(vMatriz(3, j)) - 1))
'                        xlHoja1.Range(xlHoja1.Cells(nContador, 12), xlHoja1.Cells(nContador, 12)).Borders.LineStyle = 1
'                        nContadorGrupoEstado = True
'               End If
'               If nContadorColumna = 3 Then
'                        xlHoja1.Cells(nContador, 14) = nContadorGrupo
'                        xlHoja1.Cells(nContador, 15) = Trim(Mid(vMatriz(2, j), 1, Len(vMatriz(2, j)) - 4))
'                        xlHoja1.Range(xlHoja1.Cells(nContador, 15), xlHoja1.Cells(nContador, 15 + 2)).Merge
'                        xlHoja1.Cells(nContador, 18) = Trim(Mid(vMatriz(3, j), 1, Len(vMatriz(3, j)) - 1))
'                        xlHoja1.Range(xlHoja1.Cells(nContador, 18), xlHoja1.Cells(nContador, 18)).Borders.LineStyle = 1
'                        nContadorColumna = 0
'                        nContador = nContador + 2
'                        nContadorGrupoEstado = False
'               End If
'
'             End If
'             ' ***********************Fin agregado
'
'
'               '************** Comentado por pti1 18072018 Memorandum Nº 1602-2018-GM-DI/CMACM
'                'If nFilaCan <= IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 6, 4) Then
'                   ' If Right(vMatriz(3, j), 1) <> 3 Then
'                      '  xlHoja1.Cells(nContador, 2) = j + 1
'                       ' xlHoja1.Cells(nContador, 3) = Trim(Mid(vMatriz(2, j), 1, Len(vMatriz(2, j)) - 4))
'                       ' xlHoja1.Range(xlHoja1.Cells(nContador, 3), xlHoja1.Cells(nContador, 3 + 2)).Merge
'                       ' xlHoja1.Cells(nContador, 6) = Trim(Mid(vMatriz(3, j), 1, Len(vMatriz(3, j)) - 1))
'                      '  xlHoja1.Range(xlHoja1.Cells(nContador, 6), xlHoja1.Cells(nContador, 6)).Borders.LineStyle = 1
'                      '  nContador = nContador + 2
'                      '  nFilaCan = nFilaCan + 1
'                    'End If
'                'ElseIf nFilaCan <= IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 12, 8) Then
'                    'If Right(vMatriz(3, j), 1) <> 3 Then
'                      '  xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 12, 8), 8) = j + 1
'                      '  xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 12, 8), 9) = Trim(Mid(vMatriz(2, j), 1, Len(vMatriz(2, j)) - 4))
'                       ' xlHoja1.Range(xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 12, 8), 9), xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 12, 8), 9 + 2)).Merge
'                       ' xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 12, 8), 12) = Trim(Mid(vMatriz(3, j), 1, Len(vMatriz(3, j)) - 1))
'                       ' xlHoja1.Range(xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 12, 8), 12), xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 12, 8), 12)).Borders.LineStyle = 1
'                       ' nContador = nContador + 2
'                       ' nFilaCan = nFilaCan + 1
'                   ' End If
'                'Else
'                    'If Right(vMatriz(3, j), 1) <> 3 Then
'                       ' xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 24, 16), 14) = j + 1
'                      '  xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 24, 16), 15) = Trim(Mid(vMatriz(2, j), 1, Len(vMatriz(2, j)) - 4))
'                      '  xlHoja1.Range(xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 24, 16), 15), xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 24, 16), 15 + 2)).Merge
'                      '  xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 24, 16), 18) = Trim(Mid(vMatriz(3, j), 1, Len(vMatriz(3, j)) - 1))
'                      '  xlHoja1.Range(xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 24, 16), 18), xlHoja1.Cells(nContador - IIf(cboCategoria.Text = "REQUISITOS ADICIONALES", 24, 16), 18)).Borders.LineStyle = 1
'                      '  nContador = nContador + 2
'                      '  nFilaCan = nFilaCan + 1
'                    'End If
'                'End If
'                ' Fin Comentado *****************************
'            End If
'        Next j
'        If nContadorGrupoEstado = True Then 'AGREGADO POR PTI1 18072018 Memorandum Nº 1602-2018-GM-DI/CMACM
'            nContador = nContador + 2 'AGREGADO POR PTI1 18072018 Memorandum Nº 1602-2018-GM-DI/CMACM
'        End If 'AGREGADO POR PTI1 18072018 Memorandum Nº 1602-2018-GM-DI/CMACM
'
'    Next i
'    xlHoja1.Cells(nContador + 4, 3) = "RECOMENDACIÓN Y/O OBSERVACIONES"
'
'    Dim nContadorObs As Integer
'    nContadorObs = 1
'    nContador = nContador + 4
'    Set rsCred = oCred.BuscaObsAdmCred(sCtaCod)
'
'    If Not (rsCred.EOF And rsCred.BOF) Then
'        Do While Not rsCred.EOF
'            xlHoja1.Cells(nContador + nContadorObs, 3) = nContadorObs
'            xlHoja1.Cells(nContador + nContadorObs, 4) = Trim(rsCred!cDescripcion)
'            xlHoja1.Range(xlHoja1.Cells(nContador + nContadorObs, 4), xlHoja1.Cells(nContador + nContadorObs, 18)).Merge
'            xlHoja1.Range(xlHoja1.Cells(nContador + nContadorObs, 3), xlHoja1.Cells(nContador + nContadorObs, 18)).Borders.LineStyle = 5
'
'            nContadorObs = nContadorObs + 1
'            rsCred.MoveNext
'
'        Loop
'        rsCred.Close
'        Set rsCred = Nothing
'    End If
'
'    Set oCred = Nothing
'    Dim psArchivoAGrabarC As String
'
'    xlHoja1.SaveAs App.Path & lsArchivo
'    psArchivoAGrabarC = App.Path & lsArchivo
'    xlsAplicacion.Visible = True
'    xlsAplicacion.Windows(1).Visible = True
'    Set xlsAplicacion = Nothing
'    Set xlsLibro = Nothing
'    Set xlHoja1 = Nothing
'Comento JOEP20190116 CP

'Agrego JOEP20190116 CP
    Dim obj As New COMNCredito.NCOMCredito
    Dim oCred As New COMDCredito.DCOMCreditos
    Dim rs As New ADODB.Recordset
    Dim rsObs As New ADODB.Recordset
    Dim i As Integer, j As Integer, IniTablas As Integer, Z As Integer, X As Integer
    Dim rsCred As New ADODB.Recordset
    Dim nContadorColumna As Integer, nContadorGrupo As Integer
    Dim nContadorGrupoEstado As Boolean
    Dim nContador As Integer, nFilaCan As Integer, nFilaAnt As Integer
    Dim rsPDNivel As New ADODB.Recordset
    Dim sItem As String
    Dim nCantConf As Integer
    Dim nDato As String
    Dim nEspCab As Integer, iTemCab As Integer, Item As Integer, iTemCabNewPag As Integer, itemNewPag As Integer, nCantLetras As Integer
    Dim B As Boolean
    Dim oDoc As cPDF
    Dim nExitNiv As Boolean
    Dim nCatg As Integer
'JOEP20190102 CP

    Set rs = obj.ObtenerDatosCredChekList(sCtaCod)
    Set oDoc = New cPDF
    IniTablas = 13
    nContador = IniTablas
    B = False
    nEspCab = 0
    iTemCab = 160
    Item = 170
    iTemCabNewPag = 50
    itemNewPag = 60
   
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - NEGOCIO"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "CheckList Nº " & sCtaCod
    oDoc.Title = "CheckList Nº " & sCtaCod

If Not oDoc.PDFCreate(App.Path & "\Spooler\CheckList" & sCtaCod & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
    Exit Sub
End If

'TIpo de Letra y Imagen
oDoc.Fonts.Add "F1", "Arial Narrow", TrueType, Normal, WinAnsiEncoding
oDoc.Fonts.Add "F2", "Arial Narrow", TrueType, Bold, WinAnsiEncoding
oDoc.LoadImageFromFile "C:\SICMACM_DEV\SICMAC_NEGOCIO_COM" & "\logo_cmacmaynas.bmp", "Logo"
'Posicion de la pagina
oDoc.NewPage A4_Horizontal
'Cabecera
oDoc.WImage 50, 50, 40, 80, "Logo"
oDoc.WTextBox 50, 50, 10, 80, "CAJA MAYNAS", "F2", 8, hCenter
oDoc.WTextBox 60, 50, 10, 90, "Administración de Créditos", "F2", 8, hLeft
oDoc.WRectangle 71, 50, 15, 780, 100, &HE0E0E0, True
oDoc.WTextBox 70, 50, 10, 750, "CONTROL DOCUMENTARIO DE CREDITOS", "F2", 12, hCenter
              
    If Not (rs.EOF And rs.BOF) Then
        nEspCab = CInt(Len(rs!cPersNombre) / 37) * 10
        oDoc.WTextBox 90, 50, 10, 100, "Cliente: ", "F1", 12, hLeft
        oDoc.WTextBox 90, 130, 10, 270, rs!cPersNombre, "F1", 10, hjustify
        oDoc.WTextBox 90, 420, 10, 500, "Cód. SBS: ", "F1", 12, hLeft
        oDoc.WTextBox 90, 470, 10, 500, IIf(IsNull(rs!cPersCodSbs), "00000000", rs!cPersCodSbs), "F1", 12, hLeft
        oDoc.WTextBox 90, 550, 10, 500, "Nº de créd: ", "F1", 12, hLeft
        oDoc.WTextBox 90, 660, 10, 500, rs!cCtaCod, "F1", 12, hLeft
        
        oDoc.WTextBox 92 + nEspCab, 50, 10, 500, "Producto: ", "F1", 12, hLeft
        oDoc.WTextBox 92 + nEspCab, 130, 10, 270, IIf(IsNull(rs!cTpoPrd), "", rs!cTpoPrd), "F1", 12, hLeft
        
        oDoc.WTextBox 92 + nEspCab, 420, 10, 500, "Usuario Analista: ", "F1", 12, hLeft
        oDoc.WTextBox 92 + nEspCab, 500, 10, 500, IIf(IsNull(rs!cUser), "", rs!cUser), "F1", 12, hLeft
        
        oDoc.WTextBox 92 + nEspCab, 550, 10, 500, "Fecha de desembolso: ", "F1", 12, hLeft
        oDoc.WTextBox 92 + nEspCab, 660, 10, 500, IIf(IsNull(rs!dVigencia), "__/__/____", rs!dVigencia), "F1", 12, hLeft
        
        oDoc.WTextBox (94 + nEspCab) + 10, 50, 10, 500, "Agencia: ", "F1", 12, hLeft
        oDoc.WTextBox (94 + nEspCab) + 10, 130, 10, 270, rs!cAgeDescripcion, "F1", 12, hLeft
        oDoc.WTextBox (94 + nEspCab) + 10, 420, 10, 500, "Monto: ", "F1", 12, hLeft
        oDoc.WTextBox (94 + nEspCab) + 10, 470, 10, 500, Format(rs!nMontoCol, gsFormatoNumeroView), "F1", 12, hLeft
        oDoc.WTextBox (94 + nEspCab) + 10, 550, 10, 500, "Modalidad del crédito: ", "F1", 12, hLeft
        oDoc.WTextBox (94 + nEspCab) + 10, 660, 10, 500, rs!cModalidad, "F1", 12, hLeft
    End If
        
    oDoc.WRectangle 141, 50, 15, 780, 100, &HE0E0E0, True
    oDoc.WTextBox 140, 50, 10, 750, "INFORMACIÓN EN LOS FILES DEL CRÉDITO", "F2", 12, hCenter
        
For i = 0 To cboCategoria.ListCount - 1
    cboCategoria.ListIndex = i
        If Item > 500 And B = False Then
            oDoc.NewPage A4_Horizontal
            Item = 60
            iTemCab = 50
        End If
            nDato = Trim(Right(cboCategoria.Text, 9))
            sItem = Mid(nDato, 1, (InStr(nDato, "-")) - 1)
            nCantConf = Mid(nDato, (InStr(nDato, "-")) + 1, 6)
            Set rsPDNivel = obj.ListaCredAdmNiveles(sTpoCateg, sTpoProd, sItem, nCantConf)
            
            If Not (rsPDNivel.BOF And rsPDNivel.EOF) Then
                'If ValidaExisteItem(cmbNiveles.ListIndex, 2) = True Then 'Comento JOEP20190305 Mejora CP
                    nExitNiv = True
                'End If 'Comento JOEP20190305 Mejora CP
            Else
                If ValidaExisteItem(cboCategoria.ListIndex, 1) = True Then
                    nExitNiv = False
                    oDoc.WTextBox iTemCab, 50, 10, 500, i + 1 & " " & Mid(cboCategoria.Text, 1, InStr(cboCategoria.Text, "-") - 3), "F2", 10, hLeft
                    nContador = nContador + 2
                    nFilaAnt = nContador
                    nFilaCan = 1
                End If
            End If
                
            If nExitNiv = True Then
                For Z = 0 To cmbNiveles.ListCount - 1
                    cmbNiveles.ListIndex = Z
                    If ValidaExisteItem(cmbNiveles.ListIndex, 2) = True Then
                        oDoc.WTextBox iTemCab, 50, 20, 500, i + 1 & " " & Mid(cmbNiveles.Text, 1, InStr(cmbNiveles.Text, ".") - 3), "F2", 10, hLeft
                        nContador = nContador + 2
                        nFilaAnt = nContador
                        nFilaCan = 1
                    End If
                    
                    nContadorGrupoEstado = False
                    nContadorGrupo = 0
                    nContadorColumna = 0
                    For X = 0 To UBound(vMatriz, 2)
                        If Trim(cmbNiveles.ItemData(cmbNiveles.ListIndex)) = Trim(vMatriz(1, X)) Then
                            If Item > 500 And B = False Then
                                oDoc.NewPage A4_Horizontal
                                Item = 60
                                iTemCab = 50
                            End If
                                If Right(vMatriz(3, X), 1) <> 3 Then
                                    nContadorColumna = nContadorColumna + 1
                                    nContadorGrupo = nContadorGrupo + 1
                                    If nContadorColumna = 1 Then
                                        oDoc.WTextBox Item, 50, 10, 10, nContadorGrupo & " ", "F1", 8, hLeft
                                        oDoc.WTextBox Item, 60, 10, 230, Trim(Mid(vMatriz(2, X), 1, Len(vMatriz(2, X)) - 4)), "F1", 8, hjustify
                                        oDoc.WTextBox Item, 295, 10, 250, Trim(Mid(vMatriz(3, X), 1, Len(vMatriz(3, X)) - 1)), "F2", 8, hLeft
                                        nCantLetras = Len(Trim(Mid(vMatriz(2, X), 1, Len(vMatriz(2, X)) - 4)))
                                        nContadorGrupoEstado = True
                                    End If
                                    If nContadorColumna = 2 Then
                                        oDoc.WTextBox Item, 310, 10, 10, nContadorGrupo & " ", "F1", 8, hLeft
                                        oDoc.WTextBox Item, 320, 10, 230, Trim(Mid(vMatriz(2, X), 1, Len(vMatriz(2, X)) - 4)), "F1", 8, hjustify
                                        oDoc.WTextBox Item, 555, 10, 250, Trim(Mid(vMatriz(3, X), 1, Len(vMatriz(3, X)) - 1)), "F2", 8, hLeft
                                        If Len(Trim(Mid(vMatriz(2, X), 1, Len(vMatriz(2, X)) - 4))) > nCantLetras Then
                                            nCantLetras = Len(Trim(Mid(vMatriz(2, X), 1, Len(vMatriz(2, X)) - 4)))
                                        End If
                                        nContadorGrupoEstado = True
                                    End If
                                    If nContadorColumna = 3 Then
                                        oDoc.WTextBox Item, 570, 10, 10, nContadorGrupo & " ", "F1", 8, hLeft
                                        oDoc.WTextBox Item, 580, 10, 230, Trim(Mid(vMatriz(2, X), 1, Len(vMatriz(2, X)) - 4)), "F1", 8, hjustify
                                        oDoc.WTextBox Item, 695, 10, 250, Trim(Mid(vMatriz(3, X), 1, Len(vMatriz(3, X)) - 1)), "F2", 8, hCenter
                                        If Len(Trim(Mid(vMatriz(2, X), 1, Len(vMatriz(2, X)) - 4))) > nCantLetras Then
                                            nCantLetras = Len(Trim(Mid(vMatriz(2, X), 1, Len(vMatriz(2, X)) - 4)))
                                        End If
                                        nContadorColumna = 0
                                        nContador = nContador + 2
                                        nContadorGrupoEstado = False
                                        Item = Item + IIf(nCantLetras > 45, IIf(CInt(nCantLetras / 45) = 1, 2, CInt(nCantLetras / 45)) * 14, 20)
                                        iTemCab = (Item - 10)
                                        nCantLetras = 0
                                    End If
                                End If
                        End If
                    Next X
                    If nContadorGrupoEstado = True Then
                        nContador = nContador + 2
                        If nContadorColumna <> 0 Then
                            Item = Item + IIf(nCantLetras > 45, IIf(CInt(nCantLetras / 45) = 1, 2, CInt(nCantLetras / 45)) * 10 + 10, 20) 'JOEP20190102 CP
                        End If
                        iTemCab = (Item - 10)
                        B = False
                    End If
                Next Z
            Else
                nContadorGrupoEstado = False
                nContadorGrupo = 0
                nContadorColumna = 0
                For j = 0 To UBound(vMatriz, 2)
                    If Trim(cboCategoria.ItemData(cboCategoria.ListIndex)) = Trim(vMatriz(1, j)) Then
                        If Item > 500 And B = False Then
                            oDoc.NewPage A4_Horizontal
                            Item = 60
                            iTemCab = 50
                        End If
                            If Right(vMatriz(3, j), 1) <> 3 Then
                                nContadorColumna = nContadorColumna + 1
                                nContadorGrupo = nContadorGrupo + 1
                                If nContadorColumna = 1 Then
                                    oDoc.WTextBox Item, 50, 10, 10, nContadorGrupo & " ", "F1", 8, hLeft
                                    oDoc.WTextBox Item, 60, 10, 230, Trim(Mid(vMatriz(2, j), 1, Len(vMatriz(2, j)) - 4)), "F1", 8, hjustify
                                    oDoc.WTextBox Item, 295, 10, 250, Trim(Mid(vMatriz(3, j), 1, Len(vMatriz(3, j)) - 1)), "F2", 8, hLeft
                                    nCantLetras = Len(Trim(Mid(vMatriz(2, j), 1, Len(vMatriz(2, j)) - 4)))
                                    nContadorGrupoEstado = True
                                End If
                                If nContadorColumna = 2 Then
                                    oDoc.WTextBox Item, 310, 10, 10, nContadorGrupo & " ", "F1", 8, hLeft
                                    oDoc.WTextBox Item, 320, 10, 230, Trim(Mid(vMatriz(2, j), 1, Len(vMatriz(2, j)) - 4)), "F1", 8, hjustify
                                    oDoc.WTextBox Item, 555, 10, 250, Trim(Mid(vMatriz(3, j), 1, Len(vMatriz(3, j)) - 1)), "F2", 8, hLeft
                                    If Len(Trim(Mid(vMatriz(2, j), 1, Len(vMatriz(2, j)) - 4))) > nCantLetras Then
                                        nCantLetras = Len(Trim(Mid(vMatriz(2, j), 1, Len(vMatriz(2, j)) - 4)))
                                    End If
                                    nContadorGrupoEstado = True
                                End If
                                If nContadorColumna = 3 Then
                                    oDoc.WTextBox Item, 570, 10, 10, nContadorGrupo & " ", "F1", 8, hLeft
                                    oDoc.WTextBox Item, 580, 10, 230, Trim(Mid(vMatriz(2, j), 1, Len(vMatriz(2, j)) - 4)), "F1", 8, hjustify
                                    oDoc.WTextBox Item, 695, 10, 250, Trim(Mid(vMatriz(3, j), 1, Len(vMatriz(3, j)) - 1)), "F2", 8, hCenter
                                    If Len(Trim(Mid(vMatriz(2, j), 1, Len(vMatriz(2, j)) - 4))) > nCantLetras Then
                                        nCantLetras = Len(Trim(Mid(vMatriz(2, j), 1, Len(vMatriz(2, j)) - 4)))
                                    End If
                                    nContadorColumna = 0
                                    nContador = nContador + 2
                                    nContadorGrupoEstado = False
                                    Item = Item + IIf(nCantLetras > 45, IIf(CInt(nCantLetras / 45) = 1, 2, CInt(nCantLetras / 50)) * 14, 20)  'JOEP20190102 CP
                                    iTemCab = (Item - 10)
                                    nCantLetras = 0
                                End If
                            End If
                    End If
                Next j
                
                If nContadorGrupoEstado = True Then
                    nContador = nContador + 2
                    If nContadorColumna <> 0 Then
                        Item = Item + IIf(nCantLetras > 45, IIf(CInt(nCantLetras / 45) = 1, 2, CInt(nCantLetras / 45)) * 10 + 10, 20) 'JOEP20190102 CP
                    End If
                    iTemCab = (Item - 10)
                    B = False
                End If
            End If
Next i
    
'************************************** Observaciones *****************************************************
    oDoc.WTextBox iTemCab + 10, 50, 10, 500, "RECOMENDACIÓN Y/O OBSERVACIONES", "F2", 10, hLeft
    Item = iTemCab + 30
        
    Dim nContadorObs As Integer
    Set rsCred = oCred.BuscaObsAdmCred(sCtaCod)
        nContadorObs = 1
        nContador = nContador + 4
    If Not (rsCred.EOF And rsCred.BOF) Then
        Do While Not rsCred.EOF
            oDoc.WTextBox Item, 50, 10, 750, nContadorObs & "-  ", "F1", 12, hLeft
            oDoc.WTextBox Item, 60, 10, 750, Trim(rsCred!cDescripcion), "F1", 12, hLeft
            nContadorObs = nContadorObs + 1
            rsCred.MoveNext
            Item = Item + 12
        Loop
        rsCred.Close
        Set rsCred = Nothing
    End If
    
Set oCred = Nothing
oDoc.PDFClose
oDoc.Show
End Sub

Private Function ValidaExisteItem(ByVal pnIndex As Integer, Optional ByVal pnTpCombo As Integer) As Boolean
    Dim j As Integer
    ValidaExisteItem = False
'Combo Categoria : 1 'Combo nivel : 2
If pnTpCombo = 1 Then 'joep20190116 CP
    For j = 0 To UBound(vMatriz, 2)
        If Trim(cboCategoria.ItemData(pnIndex)) = Trim(vMatriz(1, j)) Then
            If Right(vMatriz(3, j), 1) <> 3 Then
                ValidaExisteItem = True
            End If
        End If
    Next j
Else
    'joep20190116 CP
    For j = 0 To UBound(vMatriz, 2)
        If Trim(cmbNiveles.ItemData(pnIndex)) = Trim(vMatriz(1, j)) Then
            If Right(vMatriz(3, j), 1) <> 3 Then
                ValidaExisteItem = True
            End If
        End If
    Next j
    'joep20190116 CP
End If
End Function
'JOEP20190115 CP
Private Function CP_GetValMsg() As Boolean
    Dim obj As COMDCredito.DCOMCredito
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim L As Integer
    Dim nCantNA As Integer
    Dim sItem As String
    Dim nCantConf As Integer
    Dim nDato As String

    Set obj = New COMDCredito.DCOMCredito
    
    CP_GetValMsg = True
    
    bCargIni = False
    
    For i = 1 To cboCategoria.ListCount - 1
        nCantNA = 0
        cboCategoria.ListIndex = i
        nDato = Trim(Right(cboCategoria.Text, 9))
        sItem = Mid(nDato, 1, (InStr(nDato, "-")) - 1)
        nCantConf = Mid(nDato, (InStr(nDato, "-")) + 1, 6)
        
        Set rs = obj.GetMsgCheckList(sTpoProd, sItem, nCantConf)
        If Not (rs.BOF And rs.EOF) Then
            For L = 1 To feRequisitos.rows - 1
                If Right(feRequisitos.TextMatrix(L, 2), 3) = 3 Then
                    nCantNA = nCantNA + 1
                End If
            Next L
            
            If (feRequisitos.rows - 1) = nCantNA Then
                MsgBox rs!cMsgAviso & " " & Trim(Left(cboCategoria.Text, 100)), vbInformation, "Aviso"
                CP_GetValMsg = False
                bCargIni = True
                Exit Function
            End If
        End If
    Next i
    Set obj = Nothing
    RSClose rs
End Function
'JOEP20190115 CP


