VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogReqAprobacion 
   Caption         =   "Aprobacion del Plan Anual"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13410
   Icon            =   "frmLogReqAprobacion.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   13410
   Begin VB.CommandButton cmdaprobar 
      Caption         =   "Exportar "
      Height          =   375
      Index           =   2
      Left            =   9720
      TabIndex        =   13
      Top             =   6000
      Width           =   1695
   End
   Begin VB.TextBox txtconsolidado 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7440
      TabIndex        =   10
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton cmdver 
      Caption         =   "Ver"
      Height          =   375
      Left            =   11880
      TabIndex        =   6
      Top             =   75
      Width           =   1455
   End
   Begin VB.ComboBox cboPeriodo 
      Height          =   315
      ItemData        =   "frmLogReqAprobacion.frx":030A
      Left            =   840
      List            =   "frmLogReqAprobacion.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox cmbtipconsol 
      Height          =   315
      Left            =   3360
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdaprobar 
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   1
      Left            =   6360
      TabIndex        =   1
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdaprobar 
      Caption         =   "Aprobar"
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   0
      Top             =   6000
      Width           =   2055
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshAprobacion 
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   1931
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshListConsol 
      Height          =   3855
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   6800
      _Version        =   393216
      Rows            =   3
      FixedCols       =   0
      BackColorBkg    =   16777215
      FocusRect       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin Sicmact.TxtBuscar txtconsol 
      Height          =   300
      Left            =   6480
      TabIndex        =   11
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
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
      TipoBusqueda    =   2
      EnabledText     =   0   'False
   End
   Begin VB.OLE OLE1 
      Height          =   255
      Left            =   4320
      TabIndex        =   14
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Consolidado Nº"
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
      Height          =   195
      Left            =   5040
      TabIndex        =   12
      Top             =   120
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Plan Anual  Segun Consolidado Para Aprobacion"
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
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   4155
   End
   Begin VB.Label lblperiodo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
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
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Requerimiento"
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
      Height          =   195
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   1230
   End
End
Attribute VB_Name = "frmLogReqAprobacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim clsDReq As DLogRequeri
Dim clsDGnral As DLogGeneral
Dim clsDMov As DLogMov
'Pa exportar
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet





Private Sub cboPeriodo_Click()
   
   txtconsol.Enabled = True
   MshAprobacion.Clear
   MshListConsol.Clear
   trimestral MshListConsol
   aprobacion
End Sub

Private Sub cmbtipconsol_Click()
    txtconsol.Text = ""
    txtconsolidado.Text = ""
    Me.txtconsol.rs = clsDReq.CargaReqControlConsol(cboPeriodo.Text, Right(Trim(cmbtipconsol.Text), 1))
    txtconsol.Enabled = True
    MshAprobacion.Clear
    MshListConsol.Clear
    aprobacion
    If Left(cmbtipconsol.Text, 1) = "1" Then
    trimestral MshListConsol
    Else
    Mensual MshListConsol
    End If
End Sub

Private Sub cmdaprobar_Click(Index As Integer)
Dim nestado As Integer
Dim ncodigo As Integer
Dim sActualiza As String
Dim result As Integer
Set clsDMov = New DLogMov
Select Case Index
    Case 0
        'validar si ya esta aprobado 3
        ' si ya esta aprobado  entonces  no hace nada
        'si esta para aprobar entonces 2
        If cboPeriodo.Text = "" Then
            MsgBox "Seleccione el Periodo  ", vbInformation, "Selecione el Periodo"
            Exit Sub
        End If
        If cmbtipconsol.Text = "" Then
            MsgBox "Seleccione el tipo de consolidado", vbInformation, "Seleccione el Tipo de Consolidado"
            Exit Sub
        End If
        If txtconsol.Text = "" Then
            MsgBox "Seleccione Un numero de Consolidado", vbInformation, "Seleccione Un numero de Consolidado"
            Exit Sub
        End If
        nestado = clsDReq.CargaReqControlConsolEstadopoCod(cboPeriodo.Text, Right(Trim(cmbtipconsol.Text), 1), txtconsol.Text)
        'ncodigo = clsDReq.CargaReqControlConsolCodigo(cboPeriodo.Text, Right(Trim(cmbtipconsol.Text), 1))
        ncodigo = txtconsol.Text
        If nestado = 0 And ncodigo = 0 Then
            MsgBox "No Existe Consolidado para el Periodo " & cboPeriodo.Text & " y el Tipo de Requerimiento " & Left(cmbtipconsol.Text, 15), vbInformation, "No Existe Data"
            Exit Sub
        End If
        If nestado = 3 Then 'aprobado
            MsgBox "Imposible volver a Aprobar el Consolidado " & ncodigo & " del Periodo " & cboPeriodo.Text & " y el Tipo de Requerimiento " & Left(cmbtipconsol.Text, 15), vbInformation, "Este ya se Encuentra Con Aprobacion"
            Exit Sub
        ElseIf nestado = 2 Then 'Eliminado
            MsgBox "Este Consolidado " & ncodigo & " se encuentra Eliminado", vbInformation, "Consulte con su administrador del sistema"
            Exit Sub
        ElseIf nestado = 1 Then 'para aprobacion
            sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            'Cambiar el estado del consolidado a Aprobado
            If MsgBox("Desea Aprobar el Consolidado " & ncodigo & "  Para el Plan Anual del Periodo " & cboPeriodo.Text & " y Requerimiento " & Left(cmbtipconsol.Text, 15), vbQuestion + vbYesNo, "Desea Aprobar El Plan Anual ? ") = vbYes Then
                    result = clsDMov.ActualizaReqControlConsol(ncodigo, cboPeriodo.Text, Right(Trim(cmbtipconsol.Text), 1), nestado, sActualiza)
                    If result <> 0 Then
                       MsgBox "Consulte Con su Administrador del sistema"
                    ElseIf result = 0 Then
                       MsgBox "El Consolidado " & ncodigo & " del Periodo " & cboPeriodo.Text & " y el Tipo de Requerimiento " & Left(cmbtipconsol.Text, 12) & "  Se Aprobo de Manera Satisfactoria", vbInformation, "Se Aprobo de Manera Satisfactoria"
                       
                    End If
                    Set rs = clsDReq.CargaReqControlConsol(cboPeriodo.Text, Right(Trim(cmbtipconsol.Text), 1))
                    If rs.RecordCount > 0 Then
                       Set MshAprobacion.DataSource = rs
                       Else
                       MshAprobacion.Clear
                    End If
            End If
        End If
        
    Case 1
            'MsgBox "Realize la Eliminacion del Consolidado " & cboPeriodo.Text & " y Requerimiento " & Left(cmbtipconsol.Text, 15) & " Y realize los Cambios Pertinentes ", vbQuestion, " "
            Unload Me
    Case 2
            exportar
End Select

End Sub

Private Sub cmdver_Click()
Dim bvalor As Boolean


If cboPeriodo.Text = "" Then Exit Sub
If cmbtipconsol.Text = "" Then Exit Sub
If txtconsol.Text = "" Then
    MsgBox "Seleccione Un numero de Consolidado", vbInformation, "Seleccione Un numero de Consolidado"
    Exit Sub
End If

If Right(Trim(cmbtipconsol.Text), 1) = "1" Then 'regular
    Set rs = clsDReq.CargaReqConsolMensual(cboPeriodo.Text, Right(Trim(cmbtipconsol.Text), 1), True, "", "", 1, 12, "o", "3", txtconsol.Text)
Else

    Set rs = clsDReq.CargaReqConsolMensual(cboPeriodo.Text, Right(Trim(cmbtipconsol.Text), 1), True, "", "", 1, 12, "o", "1", txtconsol.Text)
End If

bvalor = False
If rs.RecordCount > 0 Then
   Set MshListConsol.DataSource = rs
   If Right(Trim(cmbtipconsol.Text), 1) = "1" Then
      trimestral MshListConsol
      Else
      Mensual MshListConsol
   End If
   bvalor = True
   Else
    MshListConsol.Clear
    If Right(Trim(cmbtipconsol.Text), 1) = "1" Then
        trimestral MshListConsol
        Else
        Mensual MshListConsol
    End If
End If

Set rs = clsDReq.CargaReqControlConsol(cboPeriodo.Text, Right(Trim(cmbtipconsol.Text), 1))
If rs.RecordCount > 0 Then
   Set MshAprobacion.DataSource = rs
   bvalor = True
   Else
   MshAprobacion.Clear
   aprobacion
End If

If bvalor = False Then
   MsgBox "No Existe Consolidado para el Periodo " & cboPeriodo.Text & " y el Tipo de Requerimiento " & Left(cmbtipconsol.Text, 15), vbInformation, "No Existe Data"
   Exit Sub
End If


End Sub

Private Sub Form_Load()
Set rs = New ADODB.Recordset
Set clsDGnral = New DLogGeneral
Set clsDReq = New DLogRequeri
'Set rs = clsDReq.CargaReqConsolMensual(cboPeriodo.Text, Right(Trim(cmbtipconsol.Text), 1), barea, scodagencia, scodarea, Trim(Right(Trim(cmbmesini.Text), 2)), Trim(Right(Trim(cmbmesfin.Text), 2)), psCategoria)
trimestral MshListConsol
Set rs = clsDGnral.CargaPeriodo
Call CargaCombo(rs, cboPeriodo)
Me.Width = 13530
cboPeriodo.ListIndex = 0
cmbtipconsol.AddItem "Regular                                           1"
cmbtipconsol.AddItem "Extemporaneo                                      2"
cmbtipconsol.ListIndex = 0
Set rs = Nothing
MshAprobacion.Cols = 5
MshAprobacion.ColWidth(0) = 1000
MshAprobacion.ColWidth(1) = 7000
End Sub
Sub aprobacion()
MshAprobacion.TextMatrix(0, 0) = "Consol.Nº"
MshAprobacion.TextMatrix(0, 1) = "Periodo - Requerimiento - Estado - Ult.Actualizacion "
End Sub
Public Sub Mensual(grilla As MSHFlexGrid)
grilla.Cols = 27
grilla.FixedRows = 2
grilla.TextMatrix(0, 0) = "Codigo de Bien"
grilla.TextMatrix(1, 0) = "Codigo de Bien"
grilla.TextMatrix(0, 1) = "Enero"
grilla.TextMatrix(0, 2) = "Enero"
grilla.TextMatrix(0, 3) = "Febrero"
grilla.TextMatrix(0, 4) = "Febrero"
grilla.TextMatrix(0, 5) = "Marzo"
grilla.TextMatrix(0, 6) = "Marzo"
grilla.TextMatrix(0, 7) = "Abril"
grilla.TextMatrix(0, 8) = "Abril"
grilla.TextMatrix(0, 9) = "Mayo"
grilla.TextMatrix(0, 10) = "Mayo"
grilla.TextMatrix(0, 11) = "Junio"
grilla.TextMatrix(0, 12) = "Junio"
grilla.TextMatrix(0, 13) = "Julio"
grilla.TextMatrix(0, 14) = "Julio"
grilla.TextMatrix(0, 15) = "Agosto"
grilla.TextMatrix(0, 16) = "Agosto"
grilla.TextMatrix(0, 17) = "Setiembre"
grilla.TextMatrix(0, 18) = "Setiembre"
grilla.TextMatrix(0, 19) = "Octubre"
grilla.TextMatrix(0, 20) = "Octubre"
grilla.TextMatrix(0, 21) = "Noviembre"
grilla.TextMatrix(0, 22) = "Noviembre"
grilla.TextMatrix(0, 23) = "Diciembre"
grilla.TextMatrix(0, 24) = "Diciembre"
grilla.TextMatrix(0, 25) = "Total"
grilla.TextMatrix(0, 26) = "Total"
grilla.MergeCells = flexMergeRestrictColumns
grilla.MergeCells = flexMergeRestrictRows
grilla.MergeRow(0) = True
grilla.MergeCol(0) = True
grilla.ColWidth(0) = 3500
grilla.TextMatrix(1, 1) = "Cant.Enero"
grilla.TextMatrix(1, 2) = "Mont.Enero"
grilla.TextMatrix(1, 3) = "Cant.Febrero"
grilla.TextMatrix(1, 4) = "Mont.Febrero"
grilla.TextMatrix(1, 5) = "Cant.Marzo"
grilla.TextMatrix(1, 6) = "Mont.Marzo"
grilla.TextMatrix(1, 7) = "Cant.Abril"
grilla.TextMatrix(1, 8) = "Mont.Abril"
grilla.TextMatrix(1, 9) = "Cant.Mayo"
grilla.TextMatrix(1, 10) = "Mont.Mayo"
grilla.TextMatrix(1, 11) = "Cant.Junio"
grilla.TextMatrix(1, 12) = "Mont.Junio"
grilla.TextMatrix(1, 13) = "Cant.Julio"
grilla.TextMatrix(1, 14) = "Mont.Julio"
grilla.TextMatrix(1, 15) = "Cant.Agosto"
grilla.TextMatrix(1, 16) = "Mont.Agosto"
grilla.TextMatrix(1, 17) = "Cant.Setiembre"
grilla.TextMatrix(1, 18) = "Mont.Setiembre"
grilla.TextMatrix(1, 19) = "Cant.Octubre"
grilla.TextMatrix(1, 20) = "Mont.Octubre"
grilla.TextMatrix(1, 21) = "Cant.Noviembre"
grilla.TextMatrix(1, 22) = "Mont.Noviembre"
grilla.TextMatrix(1, 23) = "Cant.Diciembre"
grilla.TextMatrix(1, 24) = "Mont.Diciembre"
grilla.TextMatrix(1, 25) = "Cant.Total"
grilla.TextMatrix(1, 26) = "Mont.Total"
grilla.ColAlignment(1) = flexAlignRightCenter
grilla.ColAlignment(2) = flexAlignRightCenter
grilla.ColAlignment(3) = flexAlignRightCenter
grilla.ColAlignment(4) = flexAlignRightCenter
grilla.ColAlignment(5) = flexAlignRightCenter
grilla.ColAlignment(6) = flexAlignRightCenter
grilla.ColAlignment(7) = flexAlignRightCenter
grilla.ColAlignment(8) = flexAlignRightCenter
grilla.ColAlignment(9) = flexAlignRightCenter
grilla.ColAlignment(11) = flexAlignRightCenter
grilla.ColAlignment(12) = flexAlignRightCenter
grilla.ColAlignment(13) = flexAlignRightCenter
grilla.ColAlignment(14) = flexAlignRightCenter
grilla.ColAlignmentFixed(-1) = flexAlignCenterCenter
End Sub

Sub FormatoAprobacion()
'nLogReqPeriodo nLogControlCod nLogReqTpo  nLogConsolEstado cUltimaActualizacion
MshAprobacion.TextMatrix(0, 1) = "Consolidado Nº"
MshAprobacion.TextMatrix(0, 1) = "Periodo "
MshAprobacion.TextMatrix(0, 1) = "Tipo Consol"
MshAprobacion.TextMatrix(0, 1) = "Estado Consol"
MshAprobacion.TextMatrix(0, 1) = "Actualizacion"
End Sub

Private Sub txtconsol_EmiteDatos()
Me.txtconsolidado.Text = txtconsol.psDescripcion
End Sub

Private Sub txtconsol_GotFocus()
    Me.txtconsol.rs = clsDReq.CargaReqControlConsol(cboPeriodo.Text, Right(Trim(cmbtipconsol.Text), 1))
    txtconsol.Enabled = True
    MshAprobacion.Clear
    MshListConsol.Clear
    trimestral MshListConsol
End Sub

Public Sub trimestral(grilla As MSHFlexGrid)
grilla.Cols = 11
grilla.FixedRows = 2
grilla.TextMatrix(0, 0) = "Codigo de Bien"
grilla.TextMatrix(1, 0) = "Codigo de Bien"
grilla.TextMatrix(0, 1) = "Trimestre I"
grilla.TextMatrix(0, 2) = "Trimestre I"
grilla.TextMatrix(0, 3) = "Trimestre II"
grilla.TextMatrix(0, 4) = "Trimestre II"
grilla.TextMatrix(0, 5) = "Trimestre III"
grilla.TextMatrix(0, 6) = "Trimestre III"
grilla.TextMatrix(0, 7) = "Trimestre IV"
grilla.TextMatrix(0, 8) = "Trimestre IV"
grilla.TextMatrix(0, 9) = "Total Anual"
grilla.TextMatrix(0, 10) = "Total Anual"
grilla.MergeCells = flexMergeRestrictColumns
grilla.MergeCells = flexMergeRestrictRows
grilla.MergeRow(0) = True
grilla.MergeCol(0) = True
grilla.TextMatrix(1, 1) = "Cant. I"
grilla.TextMatrix(1, 2) = "Mont. I"
grilla.ColWidth(0) = 3500
grilla.TextMatrix(1, 3) = "Cant. II"
grilla.TextMatrix(1, 4) = "Mont. II"
grilla.TextMatrix(1, 5) = "Cant. III"
grilla.TextMatrix(1, 6) = "Mont. III"
grilla.TextMatrix(1, 7) = "Cant. IV"
grilla.TextMatrix(1, 8) = "Mont. IV"
grilla.TextMatrix(1, 9) = "Cant. Anual"
grilla.TextMatrix(1, 10) = "Mont. Anual"
grilla.TextMatrix(1, 1) = "Cant. I"
grilla.TextMatrix(1, 2) = "Mont. I"
grilla.ColAlignment(1) = flexAlignRightCenter
grilla.ColAlignment(2) = flexAlignRightCenter
grilla.ColAlignment(3) = flexAlignRightCenter
grilla.ColAlignment(4) = flexAlignRightCenter
grilla.ColAlignment(5) = flexAlignRightCenter
grilla.ColAlignment(6) = flexAlignRightCenter
grilla.ColAlignment(7) = flexAlignRightCenter
grilla.ColAlignment(8) = flexAlignRightCenter
grilla.ColAlignment(9) = flexAlignRightCenter
grilla.ColAlignment(10) = flexAlignRightCenter
grilla.ColAlignmentFixed(-1) = flexAlignCenterCenter
End Sub

Sub exportar()
Dim i As Long
Dim n As Long
Dim lsArchivoN As String
Dim lbLibroOpen As Boolean
Dim lsCadAnt As String
Dim lnIni As Integer
Dim j As Integer
On Error Resume Next
lsArchivoN = App.path & "\prueba.xls"
OLE1.Class = "ExcelWorkSheet"
lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
If Not lbLibroOpen Then
   Err.Clear
   'Set objExcel = CreateObject("Excel.Application")
   If Err.Number Then
      MsgBox "Can't open Excel."
   End If
   Exit Sub
End If
Set xlHoja1 = xlLibro.Worksheets(1)
ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
Dim band  As Boolean
Dim letra As String
lnIni = 0


xlHoja1.Cells(2, 1).value = "Aprobacion del Plan Anual"
xlHoja1.Cells(3, 1).value = "Periodo"
xlHoja1.Cells(3, 2).value = cboPeriodo.Text
xlHoja1.Cells(4, 1).value = "Tipo Requerimiento"
xlHoja1.Cells(4, 2).value = Left(requerimiento, 12)
xlHoja1.Cells(5, 1).value = "Consolidado "
xlHoja1.Cells(5, 2).value = "Nº:" & txtconsol.Text & " - " & txtconsolidado.Text


For n = 0 To MshListConsol.Cols - 1
    MshListConsol.Col = n
    lnIni = 0
    For i = 0 To MshListConsol.Rows - 1
            MshListConsol.Row = i
            xlHoja1.Cells(i + 9, n + 1).value = MshListConsol.Text
    Next
Next


OLE1.Class = "ExcelWorkSheet"
ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
OLE1.SourceDoc = lsArchivoN
OLE1.Verb = 1
OLE1.Action = 1
OLE1.DoVerb -1
End Sub


