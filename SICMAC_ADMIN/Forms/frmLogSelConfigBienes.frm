VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogSelConfigBienes 
   Caption         =   "Configuracion de Bienes "
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13320
   Icon            =   "frmLogSelConfigBienes.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7275
   ScaleWidth      =   13320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   7680
      TabIndex        =   15
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4680
      TabIndex        =   14
      Top             =   6840
      Width           =   1455
   End
   Begin VB.ComboBox cmbmesini 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   480
      Width           =   1575
   End
   Begin VB.ComboBox cmbmesfin 
      Height          =   315
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   480
      Width           =   1455
   End
   Begin VB.ComboBox cmbtipconsol 
      Height          =   315
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   0
      Width           =   1575
   End
   Begin VB.ComboBox cboPeriodo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   45
      Width           =   1095
   End
   Begin VB.CommandButton cmdver 
      Caption         =   "Ver"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   450
      Width           =   1455
   End
   Begin VB.TextBox txtconsolidado 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7800
      TabIndex        =   0
      Top             =   45
      Width           =   4335
   End
   Begin Sicmact.TxtBuscar txtconsol 
      Height          =   300
      Left            =   6720
      TabIndex        =   4
      Top             =   45
      Width           =   975
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshListConsol 
      Height          =   2415
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   4260
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshListDetalle 
      Height          =   3015
      Left            =   120
      TabIndex        =   16
      Top             =   3720
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   5318
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
   Begin VB.Label lblDetalle 
      AutoSize        =   -1  'True
      Caption         =   "Detalle :"
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
      TabIndex        =   13
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label lblmes1 
      AutoSize        =   -1  'True
      Caption         =   "Mes Ini :"
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
      TabIndex        =   11
      Top             =   480
      Width           =   750
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Mes Fin:"
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
      Left            =   2640
      TabIndex        =   10
      Top             =   480
      Width           =   735
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
      Left            =   2160
      TabIndex        =   7
      Top             =   45
      Width           =   1230
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
      TabIndex        =   6
      Top             =   45
      Width           =   660
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
      Left            =   5280
      TabIndex        =   5
      Top             =   120
      Width           =   1320
   End
End
Attribute VB_Name = "frmLogSelConfigBienes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim clsDReq As DLogRequeri
Dim clsDGnral As DLogGeneral
Dim clsDMov As DLogMov
Dim clsDGAdqui As DLogAdquisi
'Pa exportar
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Private Sub cboPeriodo_Click()
   txtconsol.Enabled = True
   MshListConsol.Clear
   Mensual MshListConsol
   MshListDetalle.Clear
   Format_Grilla
End Sub
Private Sub cmbmesfin_Click()
    txtconsol.Enabled = True
   MshListConsol.Clear
   Mensual MshListConsol
   MshListDetalle.Clear
   Format_Grilla
End Sub

Private Sub cmbmesini_Click()
   txtconsol.Enabled = True
   MshListConsol.Clear
   Mensual MshListConsol
   MshListDetalle.Clear
   Format_Grilla
End Sub
Private Sub cmbtipconsol_Click()
    txtconsol.Text = ""
    txtconsolidado.Text = ""
    Me.txtconsol.rs = clsDReq.CargaReqControlConsolAprobado(cboPeriodo.Text, Right(Trim(cmbtipconsol.Text), 1))
    txtconsol.Enabled = True
    'MshAprobacion.Clear
    MshListConsol.Clear
    MshListDetalle.Clear
    Format_Grilla
    aprobacion
    If Left(cmbtipconsol.Text, 1) = "1" Then
    Mensual MshListConsol
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
            
End Select

End Sub

Private Sub cmdAceptar_Click()

If txtconsol.Text = "" Then Exit Sub
If txtconsolidado.Text = "" Then Exit Sub
If cmbmesini.Text = "" Then Exit Sub
If cmbmesfin.Text = "" Then Exit Sub

If MsgBox("Se Copiaran los Items en el Proceso de seleccion " & frmLogSelSeleccionBienes.txtSeleccion.Text & " , Desea Continuar ?  ", vbInformation + vbYesNo, "Esta Seguro que Desea Continuar ?") = vbYes Then
    frmLogSelConfigBienes.Show
    
End If

Set rs = clsDReq.CargaDetalleGenerico(MshListConsol.Text, cboPeriodo.Text, Right(Trim(cmbtipconsol.Text), 1), txtconsol.Text, Right(cmbmesini.Text, 2), Right(cmbmesfin, 2))
If rs.EOF = True Then
    MsgBox "No se Pudo Cargar el detalle consulte con su Administrador del Sistema ", vbInformation, "No se pudo cargar el Detalle"
    MshListDetalle.Clear
    Format_Grilla
    lblDetalle.Caption = "Detalle : "
    Exit Sub
    Else
    'lblDetalle.Caption = "Detalle : " & MshListConsol.Text & "  " & MshListConsol.TextMatrix(MshListConsol.Row, 1)
    Set frmLogSelSeleccionBienes.fgeBienesConfig.Recordset = rs
        frmLogSelSeleccionBienes.fgeBienesConfig.AdicionaFila
        frmLogSelSeleccionBienes.fgeBienesConfig.TextMatrix(frmLogSelSeleccionBienes.fgeBienesConfig.Rows - 1, 1) = " --------------------------- "
        frmLogSelSeleccionBienes.fgeBienesConfig.TextMatrix(frmLogSelSeleccionBienes.fgeBienesConfig.Rows - 1, 2) = " --------------------- TOTAL ------------- "
        frmLogSelSeleccionBienes.fgeBienesConfig.TextMatrix(frmLogSelSeleccionBienes.fgeBienesConfig.Rows - 1, 3) = " ------------ "
        frmLogSelSeleccionBienes.fgeBienesConfig.TextMatrix(frmLogSelSeleccionBienes.fgeBienesConfig.Rows - 1, 5) = " --------------- " 'fgeBienesConfig.SumaRow(5)
        frmLogSelSeleccionBienes.fgeBienesConfig.TextMatrix(frmLogSelSeleccionBienes.fgeBienesConfig.Rows - 1, 6) = " ------------ "
        frmLogSelSeleccionBienes.fgeBienesConfig.TextMatrix(frmLogSelSeleccionBienes.fgeBienesConfig.Rows - 1, 7) = Format(frmLogSelSeleccionBienes.fgeBienesConfig.SumaRow(7), "########.00")
        frmLogSelSeleccionBienes.txtperiodo = cboPeriodo.Text
        frmLogSelSeleccionBienes.txtRequerimiento = cmbtipconsol.Text
        frmLogSelSeleccionBienes.txtconsolidado = txtconsol & " - " & txtconsolidado
        frmLogSelSeleccionBienes.txtmesini = cmbmesini.Text
        frmLogSelSeleccionBienes.txtmesfin = cmbmesfin.Text
        frmLogSelSeleccionBienes.txtCategoria = MshListConsol.Text
        frmLogSelSeleccionBienes.cmdMant(1).Enabled = True  'editar
        frmLogSelSeleccionBienes.cmdMant(2).Enabled = False  'Grabar
        frmLogSelSeleccionBienes.cmdMant(3).Enabled = False 'Cancelar
        frmLogSelSeleccionBienes.cmdEdit(0).Enabled = False
        frmLogSelSeleccionBienes.cmdEdit(1).Enabled = False
        frmLogSelSeleccionBienes.sAccionSelBienes = "N"
    'Blanquear Referencia
    'clsDGAdqui.EliminaLogSelReferencia frmLogSelSeleccionBienes.txtSeleccion
End If
Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdver_Click()
Dim bvalor As Boolean
Dim nMesIni As Integer
Dim nMesFin As Integer
nMesIni = Right(cmbmesini.Text, 2)
nMesFin = Right(cmbmesfin.Text, 2)
If cboPeriodo.Text = "" Then Exit Sub
If cmbtipconsol.Text = "" Then Exit Sub
If txtconsol.Text = "" Then
    MsgBox "Seleccione Un numero de Consolidado", vbInformation, "Seleccione Un numero de Consolidado"
    Exit Sub
End If

If Right(Trim(cmbtipconsol.Text), 1) = "1" Then 'regular
    Set rs = clsDReq.CargaReqConsolMensual(cboPeriodo.Text, Right(Trim(cmbtipconsol.Text), 1), True, "", "", nMesIni, nMesFin, "x", "1", txtconsol.Text)
Else
    Set rs = clsDReq.CargaReqConsolMensual(cboPeriodo.Text, Right(Trim(cmbtipconsol.Text), 1), True, "", "", nMesIni, nMesFin, "x", "1", txtconsol.Text)
End If
bvalor = False
If rs.RecordCount > 0 Then
   Set MshListConsol.DataSource = rs
         Mensual MshListConsol
       bvalor = True
   Else
    MshListDetalle.Clear
    Format_Grilla
    MshListConsol.Clear
    Mensual MshListConsol
End If
'Ancho_Grilla nMesIni, nMesFin

Set rs = clsDReq.CargaReqControlConsol(cboPeriodo.Text, Right(Trim(cmbtipconsol.Text), 1))
If rs.RecordCount > 0 Then
   'Set MshAprobacion.DataSource = rs
   bvalor = True
   Else
   'MshAprobacion.Clear
   'aprobacion
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
Me.Width = 13500
'Set rs = clsDReq.CargaReqConsolMensual(cboPeriodo.Text, Right(Trim(cmbtipconsol.Text), 1), barea, scodagencia, scodarea, Trim(Right(Trim(cmbmesini.Text), 2)), Trim(Right(Trim(cmbmesfin.Text), 2)), psCategoria)
trimestral MshListConsol
Set rs = clsDGnral.CargaPeriodo
Call CargaCombo(rs, cboPeriodo)
ubicar_ano frmLogSelSeleccionBienes.cboPeriodo.Text, cboPeriodo
Set rs = clsDGnral.CargaConstante(gMeses)
Call CargaCombo(rs, cmbmesini, , 1, 0)
cmbmesini.ListIndex = 0
Call CargaCombo(rs, cmbmesfin, , 1, 0)
cmbmesfin.ListIndex = 11

cmbtipconsol.AddItem "Regular " & Space(100) & Str(ReqTipoRegular)
cmbtipconsol.AddItem "Extemporaneo" & Space(100) & Str(ReqTipoExtemporaneo)
cmbtipconsol.ListIndex = 0
Set rs = Nothing

MshListDetalle.Cols = 7
MshListDetalle.ColWidth(0) = 1500
MshListDetalle.ColWidth(1) = 4000
MshListDetalle.ColWidth(2) = 1000
MshListDetalle.ColWidth(3) = 0
MshListDetalle.ColWidth(4) = 1000
MshListDetalle.ColWidth(5) = 1000
MshListDetalle.ColWidth(6) = 1000
Set clsDGAdqui = New DLogAdquisi

End Sub
Sub aprobacion()
'MshAprobacion.TextMatrix(0, 0) = "Consol.Nº"
'MshAprobacion.TextMatrix(0, 1) = "Periodo - Requerimiento - Estado - Ult.Actualizacion "
End Sub
Public Sub Mensual(grilla As MSHFlexGrid)
grilla.Cols = 28
grilla.FixedRows = 2
grilla.TextMatrix(0, 0) = "Codigo de Bien"
grilla.TextMatrix(1, 0) = "Codigo de Bien"
grilla.TextMatrix(0, 1) = "Descripcion Bien"
grilla.TextMatrix(1, 1) = "Descripcion Bien"
grilla.TextMatrix(0, 2) = "Enero"
grilla.TextMatrix(0, 3) = "Enero"
grilla.TextMatrix(0, 4) = "Febrero"
grilla.TextMatrix(0, 5) = "Febrero"
grilla.TextMatrix(0, 6) = "Marzo"
grilla.TextMatrix(0, 7) = "Marzo"
grilla.TextMatrix(0, 8) = "Abril"
grilla.TextMatrix(0, 9) = "Abril"
grilla.TextMatrix(0, 10) = "Mayo"
grilla.TextMatrix(0, 11) = "Mayo"
grilla.TextMatrix(0, 12) = "Junio"
grilla.TextMatrix(0, 13) = "Junio"
grilla.TextMatrix(0, 14) = "Julio"
grilla.TextMatrix(0, 15) = "Julio"
grilla.TextMatrix(0, 16) = "Agosto"
grilla.TextMatrix(0, 17) = "Agosto"
grilla.TextMatrix(0, 18) = "Setiembre"
grilla.TextMatrix(0, 19) = "Setiembre"
grilla.TextMatrix(0, 20) = "Octubre"
grilla.TextMatrix(0, 21) = "Octubre"
grilla.TextMatrix(0, 22) = "Noviembre"
grilla.TextMatrix(0, 23) = "Noviembre"
grilla.TextMatrix(0, 24) = "Diciembre"
grilla.TextMatrix(0, 25) = "Diciembre"
grilla.TextMatrix(0, 26) = "Total"
grilla.TextMatrix(0, 27) = "Total"
grilla.MergeCells = flexMergeRestrictColumns
grilla.MergeCells = flexMergeRestrictRows
grilla.MergeRow(0) = True
grilla.MergeCol(0) = True
grilla.MergeCol(1) = True
grilla.ColWidth(0) = 1300
grilla.ColWidth(1) = 3500
grilla.TextMatrix(1, 2) = "Cant.Enero"
grilla.TextMatrix(1, 3) = "Mont.Enero"
grilla.TextMatrix(1, 4) = "Cant.Febrero"
grilla.TextMatrix(1, 5) = "Mont.Febrero"
grilla.TextMatrix(1, 6) = "Cant.Marzo"
grilla.TextMatrix(1, 7) = "Mont.Marzo"
grilla.TextMatrix(1, 8) = "Cant.Abril"
grilla.TextMatrix(1, 9) = "Mont.Abril"
grilla.TextMatrix(1, 10) = "Cant.Mayo"
grilla.TextMatrix(1, 11) = "Mont.Mayo"
grilla.TextMatrix(1, 12) = "Cant.Junio"
grilla.TextMatrix(1, 13) = "Mont.Junio"
grilla.TextMatrix(1, 14) = "Cant.Julio"
grilla.TextMatrix(1, 15) = "Mont.Julio"
grilla.TextMatrix(1, 16) = "Cant.Agosto"
grilla.TextMatrix(1, 17) = "Mont.Agosto"
grilla.TextMatrix(1, 18) = "Cant.Setiembre"
grilla.TextMatrix(1, 19) = "Mont.Setiembre"
grilla.TextMatrix(1, 20) = "Cant.Octubre"
grilla.TextMatrix(1, 21) = "Mont.Octubre"
grilla.TextMatrix(1, 22) = "Cant.Noviembre"
grilla.TextMatrix(1, 23) = "Mont.Noviembre"
grilla.TextMatrix(1, 24) = "Cant.Diciembre"
grilla.TextMatrix(1, 25) = "Mont.Diciembre"
grilla.TextMatrix(1, 26) = "Cant.Total"
grilla.TextMatrix(1, 27) = "Mont.Total"
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


Private Sub MshListConsol_Click()
If MshListConsol.Row <= 0 Then Exit Sub

If MshListConsol.Text = "" Then Exit Sub


Set rs = clsDReq.CargaDetalleGenerico(MshListConsol.Text, cboPeriodo.Text, Right(Trim(cmbtipconsol.Text), 1), txtconsol.Text, Right(cmbmesini.Text, 2), Right(cmbmesfin, 2))
If rs.EOF = True Then
    MsgBox "No se Pudo Cargar el detalle consulte con su Administrador del Sistema ", vbInformation, "No se pudo cargar el Detalle"
    MshListDetalle.Clear
    Format_Grilla
    lblDetalle.Caption = "Detalle : "
    Exit Sub
    Else
    lblDetalle.Caption = "Detalle : " & MshListConsol.Text & "  " & MshListConsol.TextMatrix(MshListConsol.Row, 1)
    Set MshListDetalle.DataSource = rs
End If
'pintar_flex MshListConsol



End Sub
Private Sub txtconsol_EmiteDatos()
Me.txtconsolidado.Text = txtconsol.psDescripcion
End Sub

Private Sub txtconsol_GotFocus()
    Me.txtconsol.rs = clsDReq.CargaReqControlConsolAprobado(cboPeriodo.Text, Right(Trim(cmbtipconsol.Text), 1))
    txtconsol.Enabled = True
    'MshAprobacion.Clear
    MshListConsol.Clear
    Mensual MshListConsol
    MshListDetalle.Clear
    Format_Grilla
End Sub

Public Sub trimestral(grilla As MSHFlexGrid)
grilla.Cols = 11
grilla.FixedRows = 2
grilla.TextMatrix(0, 0) = "Codigo de Bien"
grilla.TextMatrix(1, 0) = "Codigo de Bien"
grilla.TextMatrix(0, 1) = "Trimestre I"
grilla.TextMatrix(0, 2) = "Trimestre I"
grilla.TextMatrix(0, 3) = "Trimestre II "
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


Sub pintar_flex(grilla As MSHFlexGrid)
        'For c = 0 To grilla.Cols - 1
        '        grilla.Col = c
        '        grilla.CellForeColor = &HFFFF00       'HFFFFFF
        '        grilla.CellBackColor = RGB(100, 200, 300)
        '        'Set grilla.CellPicture = _
        '        'LoadPicture(App.path & "\imagenes" & "\" & "Fondordeve.jpg")
        'Next
        
            'For c = 4 To 6
            'grilla.Col = c
            'grilla.Text = Format(grilla.Text, "####0.#0")
            'Next

End Sub

Sub Format_Grilla()
MshListDetalle.Cols = 6
MshListDetalle.TextMatrix(0, 0) = "Codigo de Bien"
MshListDetalle.TextMatrix(0, 1) = "Descripcion  de Bien"
MshListDetalle.TextMatrix(0, 2) = "Unidad"
MshListDetalle.TextMatrix(0, 3) = "Cantidad"
MshListDetalle.TextMatrix(0, 4) = "Monto"
MshListDetalle.TextMatrix(0, 5) = "Precio Ref"
End Sub
'Sub Ancho_Grilla(mes_Ini As Integer, mes_Fin As Integer)
'Dim A
'A = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)

'For i = 2 To MshListConsol.Cols - 3
'        MshListConsol.ColWidth(i) = 0
'Next

'For i = mes_Ini To mes_Fin * 2 Step 1
'        MshListConsol.ColWidth(i) = 1200
'Next

'End Sub
