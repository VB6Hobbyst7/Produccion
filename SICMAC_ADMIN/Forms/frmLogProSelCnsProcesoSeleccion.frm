VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogProSelCnsProcesoSeleccion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consultar Proceso de Seleccion"
   ClientHeight    =   5280
   ClientLeft      =   90
   ClientTop       =   2460
   ClientWidth     =   11655
   Icon            =   "frmLogProSelCnsProcesoSeleccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   10320
      TabIndex        =   11
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Frame FrameConsultar 
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11655
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6540
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   4800
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Frame Frame1 
         Height          =   795
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   7035
         Begin VB.ComboBox cboMes 
            Height          =   315
            Left            =   4095
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   300
            Width           =   1995
         End
         Begin VB.TextBox txtAnio 
            Height          =   315
            Left            =   6075
            MaxLength       =   4
            TabIndex        =   8
            Top             =   300
            Width           =   675
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Adquisiciones y Contrataciones para el Mes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   195
            TabIndex        =   10
            Top             =   360
            Width           =   3720
         End
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Seleccionar"
         Height          =   375
         Left            =   8640
         TabIndex        =   6
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Frame FrameOpt 
         Enabled         =   0   'False
         Height          =   795
         Left            =   7200
         TabIndex        =   2
         Top             =   0
         Width           =   4335
         Begin VB.OptionButton OptTipo 
            Caption         =   "Consultar"
            Height          =   255
            Index           =   1
            Left            =   1500
            TabIndex        =   5
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OptTipo 
            Caption         =   "No Programados"
            Height          =   255
            Index           =   2
            Left            =   2640
            TabIndex        =   3
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton OptTipo 
            Caption         =   "Plan Anual"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   1095
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
         Height          =   3855
         Left            =   120
         TabIndex        =   0
         Top             =   840
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   6800
         _Version        =   393216
         Cols            =   18
         FixedCols       =   0
         BackColorSel    =   16775645
         ForeColorSel    =   8388608
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         BandDisplay     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   18
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5880
         TabIndex        =   13
         Top             =   4860
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin VB.Image imgOK 
      Height          =   240
      Left            =   1740
      Picture         =   "frmLogProSelCnsProcesoSeleccion.frx":08CA
      Top             =   3480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgNN 
      Height          =   240
      Left            =   1440
      Picture         =   "frmLogProSelCnsProcesoSeleccion.frx":0C0C
      Top             =   3480
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmLogProSelCnsProcesoSeleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sSQL As String
Dim xEjecutar As Boolean
Dim nMes As Integer, nAnio As Integer, Rspta As Boolean, nTipo As Integer
    
Public gvnProSelNro As Integer, gvcBSGrupoCod As String, gvcMes As String, _
        gvnNro As Integer, gvcTipo As String, gvnMonto As Double, _
        gvcMoneda As String, gvcDescripcion As String, gvnProSelTpoCod As Integer, _
        gvnProSelSubTpo As Integer, gvnAnio As Integer, gvnNroProceso As Integer, _
        gvcObjeto As String, gbBandera As Boolean, gvnObjeto As Integer, gnModalidad As Boolean, _
        gvcArchivoBases As String, gvnNroPlan  As Integer
        
''********************************************************************************************
''var nuevo
'Dim RsItem As ADODB.Recordset, nItem As Integer, gnProSelTpoCod As Integer, gnProSelSubTpo As Integer
''********************************************************************************************

Public Sub Inicio(pnTipo As Integer)
    nTipo = pnTipo
    Me.Show 1
End Sub

Private Sub cboMes_Click()
    If OptTipo(0).value Then
        OptTipo_Click 0
    ElseIf OptTipo(1).value Then
        OptTipo_Click 1
    ElseIf OptTipo(2).value Then
'        txtanio.Text = Year(gdFecSis)
'        cboMes.ListIndex = Month(gdFecSis) - 1
        OptTipo_Click 2
    End If
End Sub

'Private Sub cboMoneda_Click()
'MontoTotalNuevo
'End Sub

Private Sub cmdCancelar_Click()
'    FrameNuevo.Visible = False
    FrameConsultar.Visible = True
    OptTipo(1).value = True
    MSFlex.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    On Error GoTo cmdSelect_ClickErr
    Dim i As Integer
    If OptTipo(0).value Then
        GenerarProceso
        Unload Me
    ElseIf OptTipo(1).value Then
        With MSFlex
            i = .row
            gvcMoneda = .TextMatrix(i, 11)
            gvnAnio = .TextMatrix(i, 3)
            gvnMonto = .TextMatrix(i, 12)
            gvnNro = .TextMatrix(i, 2)
            gvnProSelNro = .TextMatrix(i, 8)
            gvnObjeto = .TextMatrix(i, 9)
            gvnProSelSubTpo = .TextMatrix(i, 7)
            gvnProSelTpoCod = .TextMatrix(i, 6)
            gvcTipo = .TextMatrix(i, 5)
            gvcDescripcion = .TextMatrix(i, 10)
            gbBandera = True
            gnModalidad = .TextMatrix(i, 15)
            gvcMes = cboMes.Text
            gvcArchivoBases = .TextMatrix(i, 16)
            gvnNroPlan = .TextMatrix(i, 13)
            gvcObjeto = .TextMatrix(i, 18)
        End With
        If nTipo = 3 Then
            frmLogProSelEtapasInfo.Inicio gvnProSelNro
        Else
            Unload Me
        End If
    ElseIf OptTipo(2).value Then
        GenerarProcesoNoProgramado
        Unload Me
    End If
    Exit Sub
cmdSelect_ClickErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub Form_Load()
    CentraForm Me
    Inicializar
    txtanio = Year(gdFecSis) ' + 1
    If nTipo = 1 Then
        FrameOpt.Enabled = True
    Else
        FrameOpt.Enabled = False
    End If
    FormaFlex
    GeneraMeses
    If cboMes.ListCount > 0 Then cboMes.ListIndex = Month(gdFecSis) - 1
End Sub

Private Sub Inicializar()
    gvnProSelNro = 0: gvcBSGrupoCod = ""
    gvnNro = 0: gvcTipo = "": gvnMonto = 0
    gvcMoneda = "": gvcDescripcion = "": gvnProSelTpoCod = 0
    gvnProSelSubTpo = 0: gvnAnio = 0: gvnNroProceso = 0
    gvcObjeto = "": gbBandera = 0: gvnObjeto = 0
End Sub

Sub GeneraMeses()
Dim oConn As New DConecta, rs As New ADODB.Recordset, sSQL As String

If oConn.AbreConexion Then
   cboMes.Clear
   sSQL = "select cMes = rtrim(substring(cNomTab,1,12)) from DBComunes..TablaCod where cCodTab like 'EZ%' and len(cCodTab)=4"
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      cboMes.Clear
      Do While Not rs.EOF
         cboMes.AddItem rs!cMes
         rs.MoveNext
      Loop
   End If
End If
End Sub

Private Sub FormaFlex()
    With MSFlex
        If OptTipo(0).value Or OptTipo(2).value Then
            .Cols = 18
            .Rows = 2
            .Clear
            .TextMatrix(0, 0) = "   ":                      .ColWidth(0) = 250
            .TextMatrix(0, 1) = "Nro":                      .ColWidth(1) = 0
            .TextMatrix(0, 2) = "Item":                     .ColWidth(2) = 800
            .TextMatrix(0, 3) = "nPlanAnualAnio":           .ColWidth(3) = 0
            .TextMatrix(0, 4) = "nPlanAnualMes":            .ColWidth(4) = 0
            .TextMatrix(0, 5) = "Proceso":                  .ColWidth(5) = 3200
            .TextMatrix(0, 6) = "nProSelTpoCod":            .ColWidth(6) = 0
            .TextMatrix(0, 7) = "nProSelSubTpo":            .ColWidth(7) = 0
            .TextMatrix(0, 8) = "cBSGrupoCod":              .ColWidth(8) = 0
            .TextMatrix(0, 9) = "nObjetoCod":               .ColWidth(9) = 0
            .TextMatrix(0, 10) = "Sintesis":                .ColWidth(10) = 4600
            .TextMatrix(0, 11) = "  ":                      .ColWidth(11) = 500:
            .TextMatrix(0, 12) = "Monto":                   .ColWidth(12) = 1000
            .TextMatrix(0, 13) = "nFuenteFinCod":           .ColWidth(13) = 0
            .TextMatrix(0, 14) = "Estado":                  .ColWidth(14) = 700
            .TextMatrix(0, 15) = "nModalidad":              .ColWidth(15) = 0
            .RowHeight(1) = 280
            .TextMatrix(0, 16) = "Bases":              .ColWidth(16) = 0
            .TextMatrix(0, 17) = "Precio U":           .ColWidth(17) = 0
            '.SelectionMode = flexSelectionFree
        Else
            .Cols = 19
            .Rows = 2
            .Clear
            .SelectionMode = flexSelectionByRow
'            .MergeCells = flexMergeRestrictRows
            .TextMatrix(0, 0) = "":                         .ColWidth(0) = 0
            .TextMatrix(0, 1) = "Nro ":                     .ColWidth(1) = 0
            .TextMatrix(0, 2) = "Item":                     .ColWidth(2) = 800
            .TextMatrix(0, 3) = "nPlanAnualAnio":           .ColWidth(3) = 0
            .TextMatrix(0, 4) = "nPlanAnualMes":            .ColWidth(4) = 0
            .TextMatrix(0, 5) = "Proceso":                  .ColWidth(5) = 2200
            .TextMatrix(0, 6) = "nProSelTpoCod":            .ColWidth(6) = 0
            .TextMatrix(0, 7) = "nProSelSubTpo":            .ColWidth(7) = 0
            .TextMatrix(0, 8) = "cBSGrupoCod":              .ColWidth(8) = 0
            .TextMatrix(0, 9) = "nObjetoCod":               .ColWidth(9) = 0
            .TextMatrix(0, 10) = "Sintesis":                .ColWidth(10) = 5800
            .TextMatrix(0, 11) = "  ":                      .ColWidth(11) = 500:
            .TextMatrix(0, 12) = "Monto":                   .ColWidth(12) = 1000
            .TextMatrix(0, 13) = "nFuenteFinCod":           .ColWidth(13) = 0
            .TextMatrix(0, 14) = "Estado":                  .ColWidth(14) = 700
            .TextMatrix(0, 15) = "nModalidad":              .ColWidth(15) = 0
            .TextMatrix(0, 16) = "Bases":              .ColWidth(16) = 0
            .WordWrap = True
        End If
    End With
End Sub

Private Sub CargarDatosPlan()
    On Error GoTo CargarDatosPlanErr
    Dim ocon As DConecta, rs As ADODB.Recordset, i As Integer
    Set ocon = New DConecta
    If ocon.AbreConexion Then
        sSQL = "select  t.cProSelTpoDescripcion,d.nPlanAnualNro, d.nPlanAnualItem, d.nPlanAnualAnio, d.nPlanAnualMes, d.nProSelTpoCod, d.nProSelSubTpo, " & _
               " d.cBSGrupoCod , d.nObjetoCod, d.cSintesis, d.nMoneda, d.nValorEstimado, d.nFuenteFinCod, d.nPlanAnualEstado, nProSelNro=isnull(s.nProSelNro,0) " & _
               " from LogPlanAnualDetalle d " & _
               " inner join LogProSelTpo t on d.nProSelTpoCod = t.nProSelTpoCod " & _
               " inner join LogPlanAnual p on p.nPlanAnualNro = d.nPlanAnualNro " & _
               " left outer join LogProSelItem s on d.nPlanAnualNro = s.nPlanAnualNro and d.nPlanAnualItem = s.nPlanAnualItem " & _
               " where d.nPlanAnualMes=" & cboMes.ListIndex + 1 & " and d.nPlanAnualEstado=1  and p.nPlanAnualEstado=2 and d.nPlanAnualAnio = " & txtanio.Text
        Set rs = ocon.CargaRecordSet(sSQL)
        
        FormaFlex
        
        Do While Not rs.EOF
            i = i + 1
            InsRow MSFlex, i
            With MSFlex
                .Col = 0
                .row = i
'                .RowHeight(i) = 800
                Set .CellPicture = imgNN
                .CellPictureAlignment = 4
                .TextMatrix(i, 1) = rs!nPlanAnualNro
                .TextMatrix(i, 2) = rs!nPlanAnualItem
                .TextMatrix(i, 3) = rs!nPlanAnualAnio
                .TextMatrix(i, 4) = rs!nPlanAnualMes
                .TextMatrix(i, 5) = rs!cProSelTpoDescripcion
                .TextMatrix(i, 6) = rs!nProSelTpoCod
                .TextMatrix(i, 7) = rs!nProSelSubTpo
                .TextMatrix(i, 8) = rs!cBSGrupoCod
                .TextMatrix(i, 9) = rs!nObjetoCod
                .TextMatrix(i, 10) = rs!cSintesis
                .Col = 11
                .row = i
                .CellAlignment = 7
                .TextMatrix(i, 11) = IIf(rs!nMoneda = 1, "S/.", "$")
                .TextMatrix(i, 12) = FNumero(rs!nValorEstimado)
                .TextMatrix(i, 13) = rs!nFuenteFinCod
                .TextMatrix(i, 14) = IIf(rs!nProselNro = 0, "Pendiente", "Generado")
                .row = 1: .Col = 1
            End With
            rs.MoveNext
        Loop
        MSFlex.row = 1
        MSFlex.ColSel = 14
        ocon.CierraConexion
    End If
    Exit Sub
CargarDatosPlanErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub


Private Sub MSFlex_Click()
    Dim nCol As Integer, nTotal As Currency
    With MSFlex
        nTotal = CDbl(txtTotal.Text)
        If OptTipo(1).value Then Exit Sub
        If Val(.TextMatrix(.row, 1)) = 0 Then Exit Sub
        If .TextMatrix(.row, 8) = "0000" Then
            MsgBox "No se Puede Seleccionar este Item por no Tener Asignado ningun Grupo Pertenece", vbInformation, "Aviso"
            Exit Sub
        End If
        nCol = .Col
        .Col = 0
        If .CellPicture = imgNN Then
            Set .CellPicture = imgOK
            nTotal = nTotal + CDbl(.TextMatrix(.row, 12))
        Else
            Set .CellPicture = imgNN
            nTotal = nTotal - CDbl(.TextMatrix(.row, 12))
        End If
        .ColSel = 14
        txtTotal.Text = FNumero(nTotal)
    End With
End Sub

Private Sub CargarRequerimientoNoProgramados()
On Error GoTo CargarRequerimientoNoProgramadosErr
    Dim ocon As DConecta, rs As ADODB.Recordset, i As Integer
    Set ocon = New DConecta
    If ocon.AbreConexion Then
        'sSQL = "select cBSGrupoCod = isnull(g.cBSGrupoCod,'0000'), cBSGrupoDescripcion = isnull(g.cBSGrupoDescripcion,'SIN GRUPO DEFINIDO'), r.cBSCod, b.cBSDescripcion, r.nMoneda, Precio=sum(r.nPrecioUnitario * r.nCantidad), r.nProSelNro, nCantidad=sum(nCantidad), r.nPrecioUnitario " & _
               " from LogProSelReqDetalle r " & _
               " inner join LogProSelReq x on r.nProSelReqNro = x.nProSelReqNro " & _
               " inner join LogProSelBienesServicios b on r.cBSCod = b.cProSelBSCod " & _
               " LEFT outer join BSGrupos g on b.cBSGrupoCod = g.cBSGrupoCod " & _
               " inner join (select distinct nProSelReqNro from LogProSelAprobacion where nEstadoAprobacion=1) a on r.nProSelReqNro = a.nProSelReqNro" & _
               " Where x.nVoBoPresupuesto = 1 and x.nSubGerenciaAdm = 1 and r.nPrecioUnitario > 0 and x.nAnio = " & txtAnio.Text & " and x.nMesEje = " & cboMes.ListIndex + 1 & _
               " group by g.cBSGrupoCod, g.cBSGrupoDescripcion, r.cBSCod, b.cBSDescripcion, r.nMoneda, r.nProSelNro, r.nPrecioUnitario " & _
               " order by r.nProselNro desc"
        sSQL = "select cBSGrupoCod = isnull(g.cBSGrupoCod,'0000'), cBSGrupoDescripcion = isnull(g.cBSGrupoDescripcion,'SIN GRUPO DEFINIDO'), r.cBSCod, b.cBSDescripcion, r.nMoneda, Precio=sum(r.nPrecioUnitario * r.nCantidad), r.nProSelNro, nCantidad=sum(nCantidad), r.nPrecioUnitario " & _
               " from LogProSelReqDetalle r " & _
               " inner join LogProSelReq x on r.nProSelReqNro = x.nProSelReqNro " & _
               " inner join LogProSelBienesServicios b on r.cBSCod = b.cProSelBSCod " & _
               " LEFT outer join BSGrupos g on b.cBSGrupoCod = g.cBSGrupoCod " & _
               " Where nProselNro = 0 and x.nVoBoPresupuesto = 1 and x.nSubGerenciaAdm = 1 and r.nPrecioUnitario > 0 and x.nAnio = " & txtanio.Text & " and x.nMesEje = " & cboMes.ListIndex + 1 & _
               " and (select count(*) from LogProSelAprobacion where nEstadoAprobacion=0 and nProSelReqNro = r.nProSelReqNro) = 0 " & _
               " group by g.cBSGrupoCod, g.cBSGrupoDescripcion, r.cBSCod, b.cBSDescripcion, r.nMoneda, r.nProSelNro, r.nPrecioUnitario " & _
               " order by r.nProselNro desc"
        Set rs = ocon.CargaRecordSet(sSQL)
        
        FormaFlex
        
        Do While Not rs.EOF
            i = i + 1
            InsRow MSFlex, i
            With MSFlex
                .Col = 0
                .row = i
'                .RowHeight(i) = 800
                Set .CellPicture = imgNN
                .CellPictureAlignment = 4
                .TextMatrix(i, 1) = i
                .TextMatrix(i, 2) = i
                .TextMatrix(i, 3) = txtanio
                .TextMatrix(i, 4) = rs!cBSCod
                .TextMatrix(i, 5) = IIf(rs!nProselNro = 0, "NO DEFINIDO", "ADJUDICACION DE MENOR CUANTIA")
                .TextMatrix(i, 6) = 1
                .TextMatrix(i, 7) = 1
                .TextMatrix(i, 8) = rs!cBSGrupoCod
                .TextMatrix(i, 9) = IIf(Left(rs!cBSCod, 2) = 11, 1, 2)
                .TextMatrix(i, 10) = rs!cBSDescripcion
                .Col = 11
                .row = i
                .CellAlignment = 7
                .TextMatrix(i, 11) = IIf(rs!nMoneda = 1, "S/.", "$")
                .TextMatrix(i, 12) = FNumero(rs!Precio)
                .TextMatrix(i, 13) = rs!nCantidad
                .TextMatrix(i, 14) = IIf(rs!nProselNro = 0, "Pendiente", "Generado")
                .TextMatrix(i, 17) = rs!nPrecioUnitario
                .row = 1: .Col = 1
            End With
            rs.MoveNext
        Loop
        MSFlex.row = 1
        MSFlex.ColSel = 14
        ocon.CierraConexion
    End If
    Exit Sub
CargarRequerimientoNoProgramadosErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub CargarDatosProcesos()
    On Error GoTo CargarDatosProcesosErr
    Dim ocon As DConecta, rs As ADODB.Recordset, i As Integer
    Set ocon = New DConecta
    If ocon.AbreConexion Then
        sSQL = "select s.cArchivoBases, t.cProSelTpoDescripcion, s.nProSelNro, s.nPlanAnualNro, s.nPlanAnualAnio, " & _
                " s.nPlanAnualMes, s.nProSelTpoCod, s.nProSelSubTpo, nNroProceso, c.cConsDescripcion as cObjeto, " & _
                " s.nObjetoCod , s.nMoneda, s.nProSelMonto, s.nProSelEstado, cSintesis, nModalidadCompra " & _
                " from LogProcesoSeleccion s " & _
                " inner join LogProSelTpo t on s.nProSelTpoCod = t.nProSelTpoCod " & _
                " left outer join constante c on s.nObjetoCod=c.nConsValor and c.nConsCod = 9048" & _
                " where s.nProSelEstado > -1 and s.nPlanAnualMes=" & cboMes.ListIndex + 1 & " and nPlanAnualAnio = " & txtanio.Text
        If Not FrameOpt.Enabled Then sSQL = sSQL & " and nVoBoPresupuesto = 1 "
        sSQL = sSQL & " order by nNroProceso "
        Set rs = ocon.CargaRecordSet(sSQL)
        
        FormaFlex
        
        Do While Not rs.EOF
            i = i + 1
            InsRow MSFlex, i
            With MSFlex
                .RowHeight(i) = 800
                .TextMatrix(i, 1) = rs!nPlanAnualNro
                .TextMatrix(i, 8) = rs!nProselNro
                .TextMatrix(i, 3) = rs!nPlanAnualAnio
                .TextMatrix(i, 4) = rs!nPlanAnualMes
                .TextMatrix(i, 5) = rs!cProSelTpoDescripcion
                .TextMatrix(i, 6) = rs!nProSelTpoCod
                .TextMatrix(i, 7) = rs!nProSelSubTpo
                .TextMatrix(i, 2) = rs!nNroProceso
                .TextMatrix(i, 9) = rs!nObjetoCod
                .TextMatrix(i, 10) = rs!cSintesis
                .Col = 11
                .row = i
                .CellAlignment = 7
                .TextMatrix(i, 11) = IIf(rs!nMoneda = 1, "S/.", "$")
                .TextMatrix(i, 12) = FNumero(rs!nProSelMonto)
                .TextMatrix(i, 13) = rs!nPlanAnualNro
                .TextMatrix(i, 14) = IIf(rs!nProSelEstado = 1, "Activo", "Terminado") 'Rs!nProSelEstado '
                .TextMatrix(i, 15) = rs!nModalidadCompra
                .TextMatrix(i, 16) = rs!cArchivoBases
                .TextMatrix(i, 18) = rs!cObjeto
                .row = 1: .Col = 1
            End With
            rs.MoveNext
        Loop
        MSFlex.row = 1
        MSFlex.ColSel = 14
        ocon.CierraConexion
    End If
    Exit Sub
CargarDatosProcesosErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub MSFlex_DblClick()
'    cmdSelect_Click
End Sub

Private Sub MSFlex_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then _
        cmdSelect_Click
End Sub

Private Sub GenerarProceso()
On Error GoTo GenerarProcesoErr
    Dim ocon As DConecta, i As Integer, nProselNro As Integer, nProSelItem As Integer
    Set ocon = New DConecta
    If Not ValidaTpo Then
        MsgBox "El Item del Plan Anual no pueden Conformar un Proceso de Seleccion", vbInformation
        Exit Sub
    End If
    With MSFlex
        .Col = 0
        nProselNro = CreaProcesoSeleccion
        If nProselNro > 0 Then
            If ocon.AbreConexion Then
                Do While i < .Rows
                
                    .row = i
                    If .TextMatrix(i, 14) = "Pendiente" And .CellPicture = imgOK Then
                        nProSelItem = NroItemProseso(nProselNro)
                        sSQL = "insert into LogProSelItem(nProSelNro,nProSelItem,nPlanAnualNro,nPlanAnualItem,cBSGrupoCod,cSintesis,nMonto)" & _
                            " values(" & nProselNro & "," & nProSelItem & "," & .TextMatrix(i, 1) & "," & .TextMatrix(i, 2) & ",'" & _
                            .TextMatrix(i, 8) & "','" & .TextMatrix(i, 10) & "'," & CDbl(.TextMatrix(i, 12)) & ")"
                        ocon.Ejecutar sSQL
                        
                        gvcMoneda = MSFlex.TextMatrix(i, 11)
                        gvnAnio = MSFlex.TextMatrix(i, 3)
'                        sSQL = "update LogPlanAnualDetalle set nPlanAnualEstado=2 where nPlanAnualNro= " & .TextMatrix(i, 1) & "and nPlanAnualItem= " & .TextMatrix(i, 2)
'                        oCon.Ejecutar sSQL
                        
                        sSQL = "insert LogProSelItemBS " & _
                                "select " & nProselNro & "," & nProSelItem & ",cProSelBSCod,nCantidad, nPrecioUnitario " & _
                                "   From LogPlanAnualDetalleBS where nPlanAnualNro=" & .TextMatrix(i, 1) & " and nPlanAnualItem=" & .TextMatrix(i, 2)
                        ocon.Ejecutar sSQL
                        
                        sSQL = "insert into LogProSelEvalFactor(nFactorNro, nProSelTpoCod, nProSelSubTpo, nObjeto, cBSGrupoCod, nProSelNro, nProSelItem, nFormula, nPuntaje, nVigente) " & _
                               " select nFactorNro, nProSelTpoCod, nProSelSubTpo, nObjeto, cBSGrupoCod, " & nProselNro & "," & nProSelItem & ", nFormula, nPuntaje, nVigente from LogProSelEvalTpoFactor where nVigente=1 and nProSelTpoCod=" & Val(.TextMatrix(i, 6)) & " and nProSelSubTpo=" & Val(.TextMatrix(i, 7)) & " and cBSGrupoCod='" & .TextMatrix(i, 8) & "'"
                        ocon.Ejecutar sSQL
                        
                        sSQL = "insert into LogProSelEvalFactorRangos(nFactorNro, nProSelTpoCod, nProSelSubTpo, nObjeto, cBSGrupoCod, nProSelNro, nProSelItem, nRangoMin, nRangoMax, nPuntaje) " & _
                               " select nFactorNro, nProSelTpoCod, nProSelSubTpo, nObjeto, cBSGrupoCod, " & nProselNro & "," & nProSelItem & ", nRangoMin, nRangoMax, nPuntaje from LogProSelEvalTpoFactorRangos where nVigente=1 and nProSelTpoCod=" & Val(.TextMatrix(i, 6)) & " and nProSelSubTpo=" & Val(.TextMatrix(i, 7)) & " and cBSGrupoCod='" & .TextMatrix(i, 8) & "'"
                        ocon.Ejecutar sSQL
                                                
                    End If
                    i = i + 1
                Loop
            ocon.CierraConexion
            End If
            MsgBox "Proceso de Seleccion Creado...", vbInformation
            gbBandera = True
        End If
    End With
Exit Sub
GenerarProcesoErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Function CreaProcesoSeleccion() As Integer
Dim i As Integer, ocon As DConecta, rs As ADODB.Recordset, nProselNro As Integer, nTotal As Currency
Dim cProSelTipo As String, cProSelTipoAbrev As String, nMonto As Currency

On Error GoTo CreaProcesoSeleccionErr


Set ocon = New DConecta
    
    With MSFlex
        .Col = 0
        i = 1
        Do While i < .Rows
            .row = i
            cProSelTipo = ""
            cProSelTipoAbrev = ""
            If Not ObtenerProcesoSeleccion(.TextMatrix(i, 9), VNumero(.TextMatrix(i, 12)), cProSelTipo, cProSelTipoAbrev) Then
               MsgBox "No se puede determinar el proceso de selección..." + Space(10), vbInformation, "Aviso"
               Exit Function
            End If
            
            If Not ValidarFactores(i) Then
               If cProSelTipoAbrev <> "AMC" Then
                  MsgBox "No existen Factores de Evaluacion para el Proceso de Seleccion", vbInformation, "Aviso"
                  Exit Function
               End If
            End If
            
            If Not ValidarEtapas(i) Then
                MsgBox "No Existen Etapas para el Proceso de Seleccion", vbInformation, "Aviso"
                Exit Function
            End If
            
            If .TextMatrix(i, 14) = "Pendiente" And .CellPicture = imgOK Then
                nTotal = MontoTotal
                sSQL = "insert into LogProcesoSeleccion(nPlanAnualNro,nProSelTpoCod,nProSelSubTpo,nProSelMonto," & _
                    "nMoneda,nObjetoCod,nPlanAnualMes,nPlanAnualAnio,nTipoCambio) " & _
                    " values(" & .TextMatrix(i, 1) & "," & .TextMatrix(i, 6) & "," & .TextMatrix(i, 7) & _
                    "," & nTotal & "," & IIf(.TextMatrix(i, 11) = "S/.", 1, 2) & _
                    "," & .TextMatrix(i, 9) & "," & IIf(Len(.TextMatrix(i, 4)) = 1, .TextMatrix(i, 4), Month(gdFecSis)) & "," & .TextMatrix(i, 3) & "," & TipoCambio(1, gdFecSis) & ")"
                If ocon.AbreConexion Then
                    ocon.Ejecutar sSQL
                    Set rs = ocon.CargaRecordSet("Select nUlt=@@identity from LogProcesoSeleccion")
                    If Not rs.EOF Then
                       nProselNro = rs!nUlt
                       
                       sSQL = "INSERT INTO LogProSelEtapa (nProSelNro,nEtapaCod,nOrden,nEstado) " & _
                               " Select " & nProselNro & ",nEtapaCod,nOrden,1 from LogProSelTpoEtapa where nProSelTpoCod = " & .TextMatrix(i, 6) & " order by nOrden "
                        ocon.Ejecutar sSQL
                        
                        CreaProcesoSeleccion = nProselNro
                        gvnMonto = nTotal
                        gvnProSelNro = nProselNro
                        gvnProSelSubTpo = .TextMatrix(i, 7)
                        gvnProSelTpoCod = .TextMatrix(i, 6)
                        gvcTipo = IIf(.TextMatrix(i, 5) = "NO DEFINIDO", "ADJUDICACION DE MENOR CUANTIA", .TextMatrix(i, 5))
                        gvnNroPlan = .TextMatrix(i, 1)
                        gvcDescripcion = .TextMatrix(i, 10)
                        gvcMes = cboMes.Text
                    End If
                    ocon.CierraConexion
                End If
                Exit Function
            End If
            i = i + 1
        Loop
    End With
Exit Function
CreaProcesoSeleccionErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

Private Function MontoTotal() As Currency
On Error GoTo MontoTotalErr
    Dim i As Integer
    With MSFlex
        .Col = 0
        i = 1
        Do While i < .Rows
            .row = i
            If .TextMatrix(i, 14) = "Pendiente" And .CellPicture = imgOK Then
                MontoTotal = MontoTotal + .TextMatrix(i, 12)
            End If
            i = i + 1
        Loop
    End With
Exit Function
MontoTotalErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

Private Function NroItemProseso(pnProSelNro As Integer) As Integer
On Error GoTo MontoTotalErr
    Dim ocon As DConecta, rs As ADODB.Recordset
    Set ocon = New DConecta
    If ocon.AbreConexion Then
        sSQL = "select isnull(max(nProSelItem),0) from LogProSelItem where nProSelNro=" & pnProSelNro
        Set rs = ocon.CargaRecordSet(sSQL)
        If Not rs.EOF Then NroItemProseso = rs(0) + 1
        ocon.CierraConexion
    End If
Exit Function
MontoTotalErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

Private Function ValidarFactores(ByVal i As Integer) As Boolean
    Dim sSQL As String, ocon As DConecta, rs As ADODB.Recordset
    sSQL = "select * from LogProSelEvalTpoFactor where nVigente=1 and nProSelTpoCod=" & Val(MSFlex.TextMatrix(i, 6)) & " and nProSelSubTpo=" & Val(MSFlex.TextMatrix(i, 7)) & " and cBSGrupoCod='" & MSFlex.TextMatrix(i, 8) & "'"
    Set ocon = New DConecta
    If ocon.AbreConexion Then
        Set rs = ocon.CargaRecordSet(sSQL)
        ocon.CierraConexion
    End If
    If Not rs.EOF Then
        ValidarFactores = True
    Else
        ValidarFactores = False
    End If
    Set rs = Nothing
End Function

Private Function ValidarEtapas(ByVal i As Integer) As Boolean
    Dim sSQL As String, ocon As DConecta, rs As ADODB.Recordset
    sSQL = " Select * from LogProSelTpoEtapa where nProSelTpoCod = " & MSFlex.TextMatrix(i, 6) & " order by nOrden "
    Set ocon = New DConecta
    If ocon.AbreConexion Then
        Set rs = ocon.CargaRecordSet(sSQL)
        ocon.CierraConexion
    End If
    If Not rs.EOF Then
        ValidarEtapas = True
    Else
        ValidarEtapas = False
    End If
    Set rs = Nothing
End Function

Private Sub GenerarProcesoNoProgramado()
On Error GoTo GenerarProcesoErr
    Dim ocon As DConecta, i As Integer, nProselNro As Integer, nProSelItem As Integer, sGrupo As String ', nValorRef As Currency
    Set ocon = New DConecta
    If Not ValidaTpo Then
        MsgBox "El Item seleccionado no pueden conformar un Proceso de Seleccion", vbInformation
        Exit Sub
    End If
'    If Not ValidarFactores Then
'        MsgBox "No se han registrado los factores de evaluacion para el proceso de seleccion", vbInformation
'        Exit Sub
'    End If
'    If Not ValidarEtapas Then
'        MsgBox "No se han registrado los factores de evaluacion para el proceso de seleccion", vbInformation
'        Exit Sub
'    End If
    With MSFlex
        .Col = 0
        nProselNro = CreaProcesoSeleccion
        If nProselNro > 0 Then
            If ocon.AbreConexion Then
                Do While i < .Rows
                    .row = i
                    If .TextMatrix(i, 14) = "Pendiente" And .CellPicture = imgOK Then
                        If sGrupo <> .TextMatrix(i, 8) Then
                            nProSelItem = NroItemProseso(nProselNro)
                            sSQL = "insert into LogProSelItem(nProSelNro,nProSelItem,nPlanAnualNro,nPlanAnualItem,cBSGrupoCod,cSintesis,nMonto)" & _
                                " values(" & nProselNro & "," & nProSelItem & "," & .TextMatrix(i, 1) & "," & .TextMatrix(i, 2) & ",'" & _
                                .TextMatrix(i, 8) & "','" & .TextMatrix(i, 10) & "'," & CalcularValorRef(.TextMatrix(i, 8)) & ")"
                            ocon.Ejecutar sSQL
'                            nValorRef = 0
                            gvcMoneda = MSFlex.TextMatrix(i, 11)
                            gvnAnio = MSFlex.TextMatrix(i, 3)
                        
                            sSQL = "insert into LogProSelEvalFactor(nFactorNro, nProSelTpoCod, nProSelSubTpo, nObjeto, cBSGrupoCod, nProSelNro, nProSelItem, nFormula, nPuntaje, nVigente) " & _
                                   " select nFactorNro, nProSelTpoCod, nProSelSubTpo, nObjeto, cBSGrupoCod, " & nProselNro & "," & nProSelItem & ", nFormula, nPuntaje, nVigente from LogProSelEvalTpoFactor where nVigente=1 and nProSelTpoCod=" & Val(.TextMatrix(i, 6)) & " and nProSelSubTpo=" & Val(.TextMatrix(i, 7)) & " and cBSGrupoCod='" & .TextMatrix(i, 8) & "'"
                            ocon.Ejecutar sSQL
                            
                            sSQL = "insert into LogProSelEvalFactorRangos(nFactorNro, nProSelTpoCod, nProSelSubTpo, nObjeto, cBSGrupoCod, nProSelNro, nProSelItem, nRangoMin, nRangoMax, nPuntaje) " & _
                                   " select nFactorNro, nProSelTpoCod, nProSelSubTpo, nObjeto, cBSGrupoCod, " & nProselNro & "," & nProSelItem & ", nRangoMin, nRangoMax, nPuntaje from LogProSelEvalTpoFactorRangos where nVigente=1 and nProSelTpoCod=" & Val(.TextMatrix(i, 6)) & " and nProSelSubTpo=" & Val(.TextMatrix(i, 7)) & " and cBSGrupoCod='" & .TextMatrix(i, 8) & "'"
                            ocon.Ejecutar sSQL
                            
                            sGrupo = .TextMatrix(i, 8)
                            
                        End If
                        
'                        nValorRef = nValorRef + Val(.TextMatrix(i, 17))
                        
                        sSQL = "insert LogProSelItemBS(nProselNro,nProSelItem,cBSCod,nCantidad, nMonto) " & _
                                "values (" & nProselNro & "," & nProSelItem & ",'" & .TextMatrix(i, 4) & "'," & .TextMatrix(i, 13) & "," & Val(.TextMatrix(i, 17)) & ")"
                        ocon.Ejecutar sSQL
                        
                        sSQL = "update LogProSelReqDetalle set nProSelNro=" & nProselNro & " where nProSelNro = 0 and cBSCod='" & .TextMatrix(i, 4) & "'"
                        ocon.Ejecutar sSQL
                                                    
                    End If
                    i = i + 1
                Loop
            ocon.CierraConexion
            End If
            MsgBox "Proceso de Seleccion Creado...", vbInformation
            gbBandera = True
        End If
    End With
Exit Sub
GenerarProcesoErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub


Private Sub OptTipo_Click(Index As Integer)
    FormaFlex
    lblTotal.Visible = False
    txtTotal.Visible = False
    txtTotal.Text = "0.00"
    Select Case Index
        Case 0
            CargarDatosPlan
        Case 1
            CargarDatosProcesos
        Case 2
'            txtanio.Text = Year(gdFecSis)
'            cboMes.ListIndex = Month(gdFecSis) - 1
            CargarRequerimientoNoProgramados
            lblTotal.Visible = True
            txtTotal.Visible = True
'            Unload Me
'            FrmProceso.Show 1
'            FrameConsultar.Visible = False
'            FrameNuevo.Visible = True
'            FormaFlexItem
'            CargarMoneda
'            cboMesNuevo.ListIndex = Month(gdFecSis) - 1
'            txtAnionuevo.Text = Year(gdFecSis)
'            Set RsItem = New ADODB.Recordset
'            RsItem.Fields.Append "cProSelBSCod", adVarChar, 12, adFldMayBeNull
'            RsItem.Fields.Append "cBSDescripcion", adVarChar, 60, adFldMayBeNull
'            RsItem.Fields.Append "cBSGrupoCod", adVarChar, 4, adFldMayBeNull
'            RsItem.Fields.Append "cBSGrupoDescripcion", adVarChar, 60, adFldMayBeNull
'            RsItem.Fields.Append "nCantidad", adInteger, , adFldMayBeNull
'            RsItem.Fields.Append "nPrecio", adCurrency, , adFldMayBeNull
'            RsItem.Open
    End Select
'    MSFlex.SetFocus
End Sub

Private Sub txtanio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboMes_Click
        MSFlex.SetFocus
        Exit Sub
    End If
    KeyAscii = DigNumEnt(KeyAscii)
End Sub

Private Function ValidaTpo() As Boolean
On Error GoTo ValidaTpoErr
    Dim i As Integer, Tpo As Integer, subTpo As Integer, nMonto As Currency, nObjeto As Integer, nMoneda As String
    i = 1
    With MSFlex
        .Col = 0
        Do While i < .Rows
            .row = i
            If .TextMatrix(i, 14) = "Pendiente" And .CellPicture = imgOK Then
                If Tpo = 0 And subTpo = 0 Then
                    Tpo = .TextMatrix(i, 6)
                    subTpo = .TextMatrix(i, 7)
                    nMonto = nMonto + .TextMatrix(i, 12)
                    nObjeto = .TextMatrix(i, 9)
                    nMoneda = .TextMatrix(i, 11)
                Else
                    If nMoneda <> .TextMatrix(i, 11) And Tpo <> .TextMatrix(i, 6) Or subTpo <> .TextMatrix(i, 7) Or nObjeto <> .TextMatrix(i, 9) Then
                        ValidaTpo = False
                        Exit Function
                        'Set .CellPicture = imgNN
                    Else
                        nMonto = nMonto + .TextMatrix(i, 12)
                    End If
                End If
            End If
            i = i + 1
        Loop
        ValidaTpo = ValidaRango(Tpo, subTpo, nMonto, nObjeto, IIf(nMoneda = "S/.", 1, 2))
    End With
    Exit Function
ValidaTpoErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

'*******************************************************************************************
'******************************************Nuevo Proveso************************************
'*******************************************************************************************

'Private Sub cmdAgregar_Click()
'    Dim Rs As ADODB.Recordset, RsGrupo As ADODB.Recordset, sMsg As String, _
'        sGrupo As String, sGrupoDes   As String
'
'    txtTotal.Text = FNumero(0)
'
'    frmLogProSelBSSelector.Show 1
'
'    If frmLogProSelBSSelector.gvrs Is Nothing Then GoTo Etiqueta
'
'    Set Rs = frmLogProSelBSSelector.gvrs
'
'    If Rs.EOF Then Exit Sub
'
'    Do While Not Rs.EOF
'        Set RsGrupo = DevuelveGrupo(Rs(0))
'        If Not RsGrupo.EOF Then
'            sGrupo = RsGrupo(0)
'            sGrupoDes = RsGrupo(1)
'        Else
'            Exit Sub
'        End If
'        If sGrupo <> "0000" Then
'            RsItem.AddNew
'            RsItem.Fields(0) = Rs(0)
'            RsItem.Fields(1) = Rs(1)
'            RsItem.Fields(2) = sGrupo
'            RsItem.Fields(3) = sGrupoDes
'            RsItem.Fields(4) = 0
'            RsItem.Fields(5) = 0
'            RsItem.Update
'        Else
'            MsgBox Rs(1) & vbCrLf & " no sera tomado en cuante por no poder " & vbCrLf & " identificar el grupo al que pertenece", vbInformation
'        End If
'        Rs.MoveNext
'    Loop
'    Set Rs = Nothing
'    Set RsGrupo = Nothing
'Etiqueta:
'    cargarItem
'End Sub
'
'Private Sub cargarItem()
'    Dim sGrupoAnt As String, i As Integer
'    RsItem.Sort = "cBSGrupoCod"
'    FormaFlexItem
'    i = MSItem.Rows - 1
'    If RsItem.EOF Or RsItem.BOF Then Exit Sub
'    RsItem.MoveFirst
'    Do While Not RsItem.EOF
'        If Not EncuentraElemento(RsItem(0)) Then
'            If sGrupoAnt <> RsItem(2) Then
'                i = i + 1
'                nItem = nItem + 1
'                sGrupoAnt = RsItem(2)
'                InsRow MSItem, i
'                MSItem.Col = 0
'                MSItem.row = i
'                MSItem.CellFontSize = 10
'                MSItem.CellFontBold = True
'                MSItem.TextMatrix(i, 0) = "-"
'                MSItem.TextMatrix(i, 1) = RsItem(2)
'                MSItem.TextMatrix(i, 3) = RsItem(3)
'                MSItem.TextMatrix(i, 4) = ""
'                MSItem.TextMatrix(i, 5) = ""
'                MSItem.TextMatrix(i, 8) = 0
'                MSItem.TextMatrix(i, 7) = nItem
'            End If
'            i = i + 1
'            InsRow MSItem, i
'            MSItem.Col = 0
'            MSItem.row = i
'            MSItem.CellFontSize = 10
'            MSItem.CellFontBold = True
'            MSItem.TextMatrix(i, 0) = ""
'            MSItem.TextMatrix(i, 1) = RsItem(0)
'            MSItem.TextMatrix(i, 3) = RsItem(1)
'            MSItem.TextMatrix(i, 4) = IIf(IsNull(RsItem(4)), 0, RsItem(4))
'            MSItem.TextMatrix(i, 5) = IIf(IsNull(RsItem(5)), 0, RsItem(5))
'            MSItem.TextMatrix(i, 8) = 0
'            MSItem.TextMatrix(i, 7) = nItem
'        End If
'        GeneraMonto nItem
'        RsItem.MoveNext
'    Loop
'End Sub

'Private Sub GeneraMonto(ByVal pnProSelItem As Integer)
'On Error GoTo GeneraMontoErr
'    Dim i As Integer, nPropuesta As Currency, npapa As Integer, nRow As Integer
'    With MSItem
'        i = 1
'        nRow = .row
'        Do While i < .Rows
'            If .TextMatrix(i, 0) <> "+" And .TextMatrix(i, 0) <> "-" Then
'            If Val(.TextMatrix(i, 7)) = pnProSelItem Then
'                    nPropuesta = nPropuesta + (Val(.TextMatrix(i, 5)) * Val(.TextMatrix(i, 4)))
'                    .TextMatrix(i, 8) = Val(.TextMatrix(i, 5)) * Val(.TextMatrix(i, 4))
'                End If
'            ElseIf .TextMatrix(i, 7) = pnProSelItem Then
'                npapa = i
'            End If
'            i = i + 1
'        Loop
'        .row = npapa
'        .TextMatrix(npapa, 8) = nPropuesta
'        .row = nRow
'    End With
'    MontoTotalNuevo
'Exit Sub
'GeneraMontoErr:
'    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
'End Sub

'Private Sub MontoTotalNuevo(Optional ByVal pnBS As Integer = 0)
'    Dim i As Integer, oCon As DConecta, Rs As ADODB.Recordset, sSQL As String, _
'        nMinimo As Currency, nMaximo As Currency, nTotal As Currency, nTipoCambio As Currency
'    Set oCon = New DConecta
'    Select Case cboMoneda.ListIndex
'        Case 0
'            nTipoCambio = 1
'        Case 1
'            nTipoCambio = TipoCambio(1, gdFecSis)
'    End Select
'
'    With MSItem
'        nTotal = 0
'        Do While i < .Rows
'            If .TextMatrix(i, 0) = "+" Or .TextMatrix(i, 0) = "-" Then
'                nTotal = Val(nTotal) + Val(.TextMatrix(i, 8))
'            End If
'            i = i + 1
'        Loop
'        txtTotal.Text = FNumero(nTotal)
'    End With
'
'    If oCon.AbreConexion Then
'        sSQL = "select nBienesMin, nBienesMax, nObrasMin, nObrasMax, nServiMin, nServiMax from LogProSelTpoRangos where nProSelTpoCod = 1 and nProSelSubTpo = 1"
'        Set Rs = oCon.CargaRecordSet(sSQL)
'        If Not Rs.EOF Then
'            Select Case pnBS
'                Case 0
'                    nMinimo = Rs!nBienesMin
'                    nMaximo = Rs!nBienesMax
'                Case 1
'                    nMinimo = Rs!nObrasMin
'                    nMaximo = Rs!nObrasMax
'                Case 2
'                    nMinimo = Rs!nServiMin
'                    nMaximo = Rs!nServiMax
'            End Select
'        End If
'        Set Rs = Nothing
'        oCon.CierraConexion
'    End If
'
'    If (nTotal * nTipoCambio) < nMinimo Then
'        TxtProceso.Text = "MONTO TOTAL MENOR AL MINIMO..."
'        TxtProceso.FontBold = True
'        TxtProceso.ForeColor = &HC0&
'        txtAbreviatura.Text = ""
'        gnProSelTpoCod = 0
'        gnProSelSubTpo = 0
'    ElseIf (nTotal * nTipoCambio) > nMaximo Then
'        TxtProceso.Text = "MONTO TOTAL MAYOR AL MAXIMO..."
'        TxtProceso.FontBold = True
'        TxtProceso.ForeColor = &HC0&
'        txtAbreviatura.Text = ""
'        gnProSelTpoCod = 0
'        gnProSelSubTpo = 0
'    Else
'        TxtProceso.Text = "ADJUDICACION DE MENOR CUANTIA"
'        TxtProceso.FontBold = False
'        TxtProceso.ForeColor = &H80000008
'        txtAbreviatura.Text = "AMC"
'        gnProSelTpoCod = 1
'        gnProSelSubTpo = 1
'    End If
'End Sub

Private Function DevuelveGrupo(ByVal pcBSCod As String) As ADODB.Recordset
    On Error GoTo DevuelveGrupoErr
    Dim ocon As New DConecta, sSQL As String, rs As ADODB.Recordset
    Set ocon = New DConecta
    sSQL = "select isnull(g.cBSGrupoCod,'0000'), isnull(g.cBSGrupoDescripcion,'Sin Grupo Definido'), b.cProSelBSCod, b.cBSDescripcion " & _
           " from LogProSelBienesServicios b " & _
           " left outer join BsGrupos g on b.cBSGrupoCod = g.cBSGrupoCod" & _
           " where b.cBScod='" & pcBSCod & "'"
    If ocon.AbreConexion Then
        Set rs = ocon.CargaRecordSet(sSQL)
        Set DevuelveGrupo = rs
        Set rs = Nothing
        ocon.CierraConexion
    End If
    Exit Function
DevuelveGrupoErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

'Sub FormaFlexItem()
'MSItem.Clear
'MSItem.Rows = 2
'MSItem.RowHeight(0) = 320
'MSItem.RowHeight(1) = 8
'MSItem.ColWidth(0) = 300: MSItem.ColAlignment(0) = 4
'MSItem.ColWidth(1) = 1000:   MSItem.ColAlignment(1) = 4:  MSItem.TextMatrix(0, 1) = " Item"
'MSItem.ColWidth(2) = 0:   MSItem.ColAlignment(2) = 4:  MSItem.TextMatrix(0, 2) = " Código"
'MSItem.ColWidth(3) = 7000:  MSItem.TextMatrix(0, 3) = " Descripción"
'MSItem.ColWidth(4) = 800:  MSItem.TextMatrix(0, 4) = " Cantidad"
'MSItem.ColWidth(5) = 800:  MSItem.TextMatrix(0, 5) = " P. Uni."
'MSItem.ColWidth(6) = 0:  MSItem.TextMatrix(0, 6) = " nProSelNro"
'MSItem.ColWidth(7) = 0:  MSItem.TextMatrix(0, 7) = " nProSelItem"
'MSItem.ColWidth(8) = 800:  MSItem.TextMatrix(0, 8) = " Precio"
'End Sub

'Private Sub cmdGuardar_Click()
'On Error GoTo cmdGuardarErr
'    Dim oCon As DConecta, sSQL As String, sGrupo As String, nProselNro As Integer, _
'        nProSelItem As Integer, sGrupoDes As String, nMonto As Currency, Rs As ADODB.Recordset
'    Set oCon = New DConecta
'
'    If Not ValidaItem Then
'        MsgBox "Debe Ingresar Todas las Cantidades y Precios...", vbInformation, "Aviso"
'        Exit Sub
'    End If
'    If Val(txtTotal.Text) = 0 Then
'        MsgBox "Total Invalido...", vbInformation, "Aviso"
'        Exit Sub
'    End If
'    If RsItem.Fields.Count = 0 Then
'        MsgBox "No Existen Bienes\Servicios...", vbInformation, "Aviso"
'        Exit Sub
'    End If
'    If RsItem.BOF And RsItem.EOF Then
'        MsgBox "No Existen Bienes\Servicios...", vbInformation, "Aviso"
'        Exit Sub
'    End If
'    If gnProSelTpoCod = 0 Or gnProSelSubTpo = 0 Then
'        MsgBox "No se Puede Encontrar el Tipo de Proceso...", vbInformation, "Aviso"
'        Exit Sub
'    End If
'
'    If oCon.AbreConexion Then
'        'oCon.BeginTrans
'
'        sSQL = "insert into LogProcesoSeleccion(nProSelTpoCod,nProSelSubTpo,nProSelMonto,nMoneda,nPlanAnualMes,nPlanAnualAnio,nTipoCambio) " & _
'               " values(" & gnProSelTpoCod & "," & gnProSelSubTpo & "," & CDbl(txtTotal.Text) & "," & cboMoneda.ItemData(cboMoneda.ListIndex) & "," & cboMesNuevo.ListIndex + 1 & "," & txtAnionuevo.Text & "," & TipoCambio(1, gdFecSis) & ")"
'        oCon.Ejecutar sSQL
'
'        Set Rs = oCon.CargaRecordSet("Select nUlt=@@identity from LogProcesoSeleccion")
'        If Not Rs.EOF Then
'           nProselNro = Rs!nUlt
'
'           sSQL = "INSERT INTO LogProSelEtapa (nProSelNro,nEtapaCod,nOrden,nEstado) " & _
'                   " Select " & nProselNro & ",nEtapaCod,nOrden,1 from LogProSelTpoEtapa where nProSelTpoCod = 1 order by nOrden "
'            oCon.Ejecutar sSQL
'        End If
'
'        RsItem.MoveFirst
'        Do While Not RsItem.EOF
'            If sGrupo <> RsItem!cBSGrupoCod Then
'                sGrupo = RsItem!cBSGrupoCod
'                sGrupoDes = RsItem!cBSGrupoDescripcion
'                nMonto = DevuelveMonto(sGrupo)
'
'                sSQL = "select isnull(max(nProSelItem),0) from LogProSelItem where nProSelNro=" & nProselNro
'                Set Rs = oCon.CargaRecordSet(sSQL)
'                If Not Rs.EOF Then nProSelItem = Rs(0) + 1
'
'                sSQL = "insert into LogProSelItem(nProSelNro,nProSelItem,cBSGrupoCod,cSintesis,nMonto)" & _
'                       " values(" & nProselNro & "," & nProSelItem & ",'" & sGrupo & "','" & sGrupoDes & "'," & nMonto & ")"
'                oCon.Ejecutar sSQL
'            End If
'
'            sSQL = "insert LogProSelItemBS(nProSelNro,nProSelItem,cProSelBSCod,nCantidad, nMonto) " & _
'                   "values (" & nProselNro & "," & nProSelItem & ",'" & RsItem!cProSelBSCod & "'," & RsItem!nCantidad & "," & RsItem!nPrecio & ")"
'            oCon.Ejecutar sSQL
'
'            RsItem.MoveNext
'        Loop
'
'        MsgBox "Proceso Registrado Correctamente...", vbInformation
'
'        gvcMoneda = IIf(cboMoneda.ListIndex = 0, "S/.", "$")
'        gvnAnio = txtAnionuevo.Text
'        gvnMonto = txtTotal.Text
'        gvnProSelNro = nProselNro
'        gvnProSelSubTpo = gnProSelSubTpo
'        gvnProSelTpoCod = gnProSelTpoCod
'        gvcTipo = TxtProceso.Text
'        gbBandera = True
'        gvcMes = cboMesNuevo.Text
'
'        Limpiar
'
'        oCon.CierraConexion
'        Unload Me
'    End If
'    Exit Sub
'cmdGuardarErr:
'    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
'End Sub

'Private Function DevuelveMonto(ByVal pcBSGrupoCod As String) As Currency
'On Error GoTo DevuelveMontoErr
'    Dim i As Integer
'    With MSItem
'        Do While i < .Rows
'            If .TextMatrix(i, 1) = pcBSGrupoCod Then
'                DevuelveMonto = Val(.TextMatrix(i, 8))
'                Exit Function
'            End If
'            i = i + 1
'        Loop
'    End With
'    Exit Function
'DevuelveMontoErr:
'    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
'End Function

'Private Sub cmdQuitar_Click()
'On Error GoTo cmdQuitarErr
'    Dim i As Integer
'    If RsItem.BOF And RsItem.EOF Then Exit Sub
'    RsItem.MoveFirst
'    Do While Not RsItem.EOF
'        If RsItem!cProSelBSCod = MSItem.TextMatrix(MSItem.row, 1) Then
'            RsItem.Delete
'            RsItem.Update
'            Exit Do
'        End If
'        RsItem.MoveNext
'    Loop
'    txtTotal.Text = FNumero(0)
'    cargarItem
'    Exit Sub
'cmdQuitarErr:
'    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
'End Sub

'Private Sub Form_Load()
'CentraForm Me
'FormaFlexItem
'CargarMoneda
'GeneraMeses
'cboMes.ListIndex = Month(gdFecSis) - 1
'txtAnio.Text = Year(gdFecSis)
'Set RsItem = New ADODB.Recordset
'RsItem.Fields.Append "cProSelBSCod", adVarChar, 12, adFldMayBeNull
'RsItem.Fields.Append "cBSDescripcion", adVarChar, 60, adFldMayBeNull
'RsItem.Fields.Append "cBSGrupoCod", adVarChar, 4, adFldMayBeNull
'RsItem.Fields.Append "cBSGrupoDescripcion", adVarChar, 60, adFldMayBeNull
'RsItem.Fields.Append "nCantidad", adInteger, , adFldMayBeNull
'RsItem.Fields.Append "nPrecio", adCurrency, , adFldMayBeNull
'RsItem.Open
'End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Set frmLogCnsProcesoSeleccion = Nothing
End Sub

'Private Sub MSItem_DblClick()
'On Error GoTo MSItemErr
'    Dim i As Integer, bTipo As Boolean
'    With MSItem
'        If Trim(.TextMatrix(.row, 0)) = "-" Then
'           .TextMatrix(.row, 0) = "+"
'           i = .row + 1
'           bTipo = True
'        ElseIf Trim(.TextMatrix(.row, 0)) = "+" Then
'           .TextMatrix(.row, 0) = "-"
'           i = .row + 1
'           bTipo = False
'        Else
'            Exit Sub
'        End If
'        Do While i < .Rows
'            If Trim(.TextMatrix(i, 0)) = "+" Or Trim(.TextMatrix(i, 0)) = "-" Then
'                Exit Sub
'            End If
'
'            If bTipo Then
'                .RowHeight(i) = 0
'            Else
'                .RowHeight(i) = 260
'            End If
'            i = i + 1
'        Loop
'    End With
'Exit Sub
'MSItemErr:
'    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
'
'End Sub

'Private Sub MSItem_GotFocus()
'If txtEditItem.Visible = False Then Exit Sub
'MSItem = txtEditItem
'txtEditItem.Visible = False
'GeneraMonto Val(MSItem.TextMatrix(MSItem.row, 7))
'ActualizaCantidadPrecio MSItem.TextMatrix(MSItem.row, 1), Val(MSItem.TextMatrix(MSItem.row, 4)), Val(MSItem.TextMatrix(MSItem.row, 5))
'End Sub

'Private Sub MSItem_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then MSItem_DblClick
'    Select Case MSItem.Col
'        Case 4, 5
'            If MSItem.TextMatrix(MSItem.row, 4) = "" Then Exit Sub
'            If IsNumeric(Chr(KeyAscii)) Then _
'                EditaFlex MSItem, txtEditItem, KeyAscii
'    End Select
'End Sub
'
'Private Sub ActualizaCantidadPrecio(ByVal pcBSCod As String, ByVal pnCantidad As Long, ByVal pnPrecio As Currency)
'On Error GoTo ActualizaCantidadPrecioErr
'    If RsItem.BOF And RsItem.EOF Then Exit Sub
'    RsItem.MoveFirst
'    Do While Not RsItem.EOF
'        If RsItem!cProSelBSCod = pcBSCod Then
'            RsItem!nPrecio = pnPrecio
'            RsItem!nCantidad = pnCantidad
'            Exit Do
'        End If
'        RsItem.MoveNext
'    Loop
'    Exit Sub
'ActualizaCantidadPrecioErr:
'    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
'End Sub
'
'Sub EditaFlex(MSFlex As Control, Edt As Control, KeyAscii As Integer)
'Select Case KeyAscii
'    Case 0 To 32
'         Edt = MSFlex
'         Edt.SelStart = 1000
'    Case Else
'         Edt = Chr(KeyAscii)
'         Edt.SelStart = 1
'End Select
'Edt.Move MSFlex.Left + MSFlex.CellLeft - 15, MSFlex.Top + MSFlex.CellTop - 15, _
'         MSFlex.CellWidth, MSFlex.CellHeight
''Edt.Text = Chr(KeyAscii) ' & MSFlex
'Edt.Visible = True
'Edt.SetFocus
'End Sub
'
'Private Function EncuentraElemento(ByVal pcBSCod As String) As Boolean
'    On Error GoTo EncuentraElementoErr
'    Dim i As Integer
'    With MSItem
'        Do While i < .Rows
'            If .TextMatrix(i, 1) = pcBSCod Then
'                EncuentraElemento = True
'                Exit Function
'            End If
'            i = i + 1
'        Loop
'        EncuentraElemento = False
'    End With
'    Exit Function
'EncuentraElementoErr:
'    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
'End Function
'
'Private Sub MSItem_LeaveCell()
'If txtEditItem.Visible = False Then Exit Sub
'MSItem = txtEditItem
'txtEditItem.Visible = False
'GeneraMonto Val(MSItem.TextMatrix(MSItem.row, 7))
'ActualizaCantidadPrecio MSItem.TextMatrix(MSItem.row, 1), MSItem.TextMatrix(MSItem.row, 4), MSItem.TextMatrix(MSItem.row, 5)
'End Sub
'
'Private Sub txtEdititem_KeyDown(KeyCode As Integer, Shift As Integer)
'EditKeyCode MSItem, txtEditItem, KeyCode, Shift
'End Sub
'
'Sub EditKeyCode(MSFlex As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
'Select Case KeyCode
'    Case 27
'         Edt.Visible = False
'         MSFlex.SetFocus
'    Case 13
'         MSFlex.SetFocus
'    Case 37                     'Izquierda
'         MSFlex.SetFocus
'         DoEvents
'         If MSFlex.Col > 1 Then
'            MSFlex.Col = MSFlex.Col - 1
'         End If
'    Case 39                     'Derecha
'         MSFlex.SetFocus
'         DoEvents
'         If MSFlex.Col < MSFlex.Cols - 1 Then
'            MSFlex.Col = MSFlex.Col + 1
'         End If
'    Case 38
'         MSFlex.SetFocus
'         DoEvents
'         If MSFlex.row > MSFlex.FixedRows + 1 Then
'            MSFlex.row = MSFlex.row - 1
'         End If
'    Case 40
'         MSFlex.SetFocus
'         DoEvents
'         'If MSFlex.Row < MSFlex.FixedRows - 1 Then
'         If MSFlex.row < MSFlex.Rows - 1 Then
'            MSFlex.row = MSFlex.row + 1
'         End If
'End Select
'End Sub
'
'Private Sub txtEditItem_KeyPress(KeyAscii As Integer)
'    KeyAscii = DigNumDec(txtEditItem, KeyAscii)
'End Sub

'Private Sub CargarMoneda()
'On Error GoTo CargarMonedaErr
'    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset
'    Set oCon = New DConecta
'    If oCon.AbreConexion Then
'        sSQL = "select nConsValor, cConsDescripcion from Constante where nConsCod = 1011"
'        Set Rs = oCon.CargaRecordSet(sSQL)
'        cboMoneda.Clear
'        Do While Not Rs.EOF
'            cboMoneda.AddItem Rs!cConsDescripcion, cboMoneda.ListCount
'            cboMoneda.ItemData(cboMoneda.ListCount - 1) = Rs!nConsValor
'            Rs.MoveNext
'        Loop
'        oCon.CierraConexion
'        cboMoneda.ListIndex = 0
'    End If
'    Exit Sub
'CargarMonedaErr:
'    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
'End Sub

'Sub GeneraMeses()
'Dim oConn As New DConecta, Rs As New ADODB.Recordset, sSQL As String
'
'If oConn.AbreConexion Then
'   cboMes.Clear
'   sSQL = "select cMes = rtrim(substring(cNomTab,1,12)) from DBComunes..TablaCod where cCodTab like 'EZ%' and len(cCodTab)=4"
'   Set Rs = oConn.CargaRecordSet(sSQL)
'   If Not Rs.EOF Then
'      Do While Not Rs.EOF
'         cboMes.AddItem Rs!cMes
'         Rs.MoveNext
'      Loop
'   End If
'   cboMes.ListIndex = 0
'End If
'End Sub

'Private Sub Limpiar()
'    Do While Not RsItem.EOF
'        RsItem.Delete
'    Loop
'    txtAbreviatura.Text = ""
'    txtAnio.Text = Year(gdFecSis)
'    TxtProceso.Text = ""
'    txtTotal.Text = 0
'    FormaFlexItem
'End Sub

'Private Function ValidaItem() As Boolean
'On Error GoTo ValidaItemErr
'    Dim i As Integer
'    With MSItem
'        i = 1
'        Do While i < .Rows
'            If Len(Trim(.TextMatrix(i, 1))) = 10 And (Val(.TextMatrix(i, 4)) = 0 Or Val(.TextMatrix(i, 5)) = 0) Then
'                ValidaItem = False
'                Exit Function
'            End If
'            i = i + 1
'        Loop
'    End With
'    ValidaItem = True
'    Exit Function
'ValidaItemErr:
'    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
'End Function

Private Function CalcularValorRef(ByVal pcGrupoCod As String)
On Error GoTo CalcularValorRefErr
    Dim i As Integer
    With MSFlex
        Do While i < .Rows
            .row = i
            If .TextMatrix(i, 14) = "Pendiente" And .CellPicture = imgOK And .TextMatrix(i, 8) = pcGrupoCod Then
                CalcularValorRef = CalcularValorRef + CDbl(.TextMatrix(i, 12))
            End If
            i = i + 1
        Loop
    End With
Exit Function
CalcularValorRefErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function
