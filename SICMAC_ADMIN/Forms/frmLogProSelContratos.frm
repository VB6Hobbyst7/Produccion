VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogProSelContrato 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Contrato"
   ClientHeight    =   5445
   ClientLeft      =   450
   ClientTop       =   2535
   ClientWidth     =   10950
   Icon            =   "frmLogProSelContratos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   10950
   Begin VB.Frame Frame2 
      Caption         =   "Datos del Proceso de Seleccion"
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
      Height          =   1530
      Left            =   120
      TabIndex        =   10
      Top             =   90
      Width           =   10695
      Begin VB.TextBox TxtDescripcion 
         Appearance      =   0  'Flat
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
         Height          =   495
         Left            =   1680
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   950
         Width           =   8835
      End
      Begin VB.TextBox TxtMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   300
         Left            =   9240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   630
         Width           =   1260
      End
      Begin VB.TextBox TxtTipo 
         Appearance      =   0  'Flat
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
         Height          =   300
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   630
         Width           =   6255
      End
      Begin VB.TextBox txtanio 
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   300
         Width           =   3255
      End
      Begin VB.CommandButton CmdConsultarProceso 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2670
         TabIndex        =   12
         Top             =   310
         Width           =   350
      End
      Begin VB.TextBox txtObjeto 
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   300
         Width           =   1860
      End
      Begin VB.TextBox TxtProSelNro 
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   1680
         TabIndex        =   14
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   1020
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Monto"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8040
         TabIndex        =   23
         Top             =   690
         Width           =   450
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Proceso Selección"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   700
         Width           =   1335
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Nº Proceso"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   345
         Width           =   810
      End
      Begin VB.Label LblMoneda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   300
         Left            =   8640
         TabIndex        =   20
         Top             =   630
         Width           =   580
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Ejecución"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3120
         TabIndex        =   19
         Top             =   360
         Width           =   1425
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Objeto"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8040
         TabIndex        =   18
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   9660
      TabIndex        =   6
      Top             =   4980
      Width           =   1155
   End
   Begin VB.CommandButton cmdContrato 
      Caption         =   "Generar Contrato"
      Height          =   375
      Left            =   7380
      TabIndex        =   5
      Top             =   4980
      Width           =   2235
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ganador"
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   4020
      Width           =   10695
      Begin VB.TextBox txtPersCod 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtPersona 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2220
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   5565
      End
      Begin VB.TextBox txtMontoGanador 
         Height          =   315
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label lblmonedaganador 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S/."
         Height          =   315
         Left            =   8880
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Propuesta"
         Height          =   195
         Left            =   7860
         TabIndex        =   9
         Top             =   420
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Postor"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   420
         Width           =   450
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSItem 
      Height          =   2340
      Left            =   120
      TabIndex        =   0
      Top             =   1620
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   4128
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   8
      FixedCols       =   0
      ForeColorFixed  =   -2147483646
      BackColorSel    =   16773857
      ForeColorSel    =   -2147483635
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      GridColorUnpopulated=   -2147483633
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
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
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
End
Attribute VB_Name = "frmLogProSelContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String
Dim nMes As Integer, nAnio As Integer
Dim gnProSelNro As Integer, gcBSGrupoCod As String

Private Sub CmdConsultarProceso_Click()
On Error GoTo msflex_clckErr
    frmLogProSelCnsProcesoSeleccion.Inicio 2
    With frmLogProSelCnsProcesoSeleccion
        If .gbBandera Then
            gnProSelNro = .gvnProSelNro
            gcBSGrupoCod = .gvcBSGrupoCod
            TxtProSelNro.Text = .gvnNro
            TxtTipo.Text = .gvcTipo
            TxtMonto.Text = Format(.gvnMonto, "###,###.00")
            LblMoneda.Caption = .gvcMoneda
            TxtDescripcion.Text = .gvcDescripcion
            
            GeneraDetalleItem gnProSelNro
       End If
    End With
Exit Sub
msflex_clckErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Sub GeneraDetalleItem(vProSelNro As Integer)
Dim oConn As New DConecta, rs As New ADODB.Recordset, i As Integer, nSuma As Currency
Dim sSQL As String, sGrupo As String

sSQL = ""
nSuma = 0
FormaFlexItem

If oConn.AbreConexion Then

    sSQL = "select v.nProSelNro, v.nProSelItem, b.cBSGrupoDescripcion, b.cBSGrupoCod,x.cBSCod, y.cBSDescripcion, x.nCantidad, v.nMonto " & _
            "from LogProSelItem v " & _
            "inner join BSGrupos b on v.cBSGrupoCod = b.cBSGrupoCod " & _
            "inner join LogProSelItemBS x on v.nProSelNro = x.nProSelNro and v.nProSelItem = x.nProSelItem " & _
            "inner join LogProSelBienesServicios y on x.cBSCod = y.cProSelBSCod " & _
            "where v.nProSelNro = " & vProSelNro & " order by v.nProSelItem, b.cBSGrupoDescripcion "

   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
        If sGrupo <> rs!nProSelItem Then
         sGrupo = rs!nProSelItem
         i = i + 1
         InsRow MSItem, i
         MSItem.Col = 0
         MSItem.row = i
         MSItem.CellFontSize = 10
         MSItem.CellFontBold = True
         MSItem.TextMatrix(i, 0) = "+"
         MSItem.TextMatrix(i, 1) = rs!nProSelItem
         MSItem.TextMatrix(i, 3) = rs!cBSGrupoDescripcion
         MSItem.TextMatrix(i, 4) = ""
         MSItem.TextMatrix(i, 5) = rs!nProselNro
         MSItem.TextMatrix(i, 6) = rs!nProSelItem
         MSItem.TextMatrix(i, 7) = ""
         MSItem.TextMatrix(i, 8) = FNumero(rs!nMonto)
         MSItem.TextMatrix(i, 9) = 0
        End If
        i = i + 1
        InsRow MSItem, i
        MSItem.RowHeight(i) = 0
        MSItem.TextMatrix(i, 1) = rs!cBSCod
        MSItem.TextMatrix(i, 3) = rs!cBSDescripcion
        MSItem.TextMatrix(i, 4) = rs!nCantidad
        MSItem.TextMatrix(i, 5) = rs!nProselNro
        MSItem.TextMatrix(i, 6) = rs!nProSelItem
        MSItem.TextMatrix(i, 7) = ""
        MSItem.TextMatrix(i, 8) = 0
        MSItem.TextMatrix(i, 9) = 0
        rs.MoveNext
      Loop
   End If
   MSItem.SetFocus
End If
End Sub


Private Sub cmdContrato_Click()
'    If Val(MSFlex.TextMatrix(MSFlex.Row, 4)) = 0 Then Exit Sub
    'If MSPos.TextMatrix(MSPos.Row, 1) = "" Then Exit Sub
    If txtPersCod.Text = "" Then Exit Sub
    'If MSPos.TextMatrix(MSPos.Row, 2) = "" Then Exit Sub
    If txtPersona.Text = "" Then Exit Sub
    'If Val(MSPos.TextMatrix(MSPos.Row, 3)) = 0 Then Exit Sub
    If TxtMonto.Text = "" Then Exit Sub
'    If Val(MSFlex.TextMatrix(MSFlex.Row, 1)) = 0 Then Exit Sub
    frmLogProSelContratoDet.Contrato gnProSelNro, txtPersCod.Text, txtPersona.Text, txtMontoGanador.Text, Val(MSItem.TextMatrix(MSItem.row, 6))
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
FormaFlexItem
txtanio.Text = Year(gdFecSis)
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmLogProSelContrato = Nothing
End Sub

Private Sub TxtProSelNro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txtanio.Text) > 0 Then
            ConsultarProcesoNro Val(TxtProSelNro.Text), Val(txtanio.Text)
            Exit Sub
        Else
            txtanio.SetFocus
        End If
    End If
    KeyAscii = DigNumEnt(KeyAscii)
End Sub

Private Sub ConsultarProcesoNro(ByVal pnNro As Integer, ByVal pnAnio As Integer)
    On Error GoTo ConsultarProcesoNroErr
    Dim oCon As DConecta, sSQL As String, rs As ADODB.Recordset
    Set oCon = New DConecta
    sSQL = "select t.cProSelTpoDescripcion, s.nProSelNro, s.nPlanAnualNro, s.nPlanAnualAnio, " & _
            "s.nPlanAnualMes, s.nProSelTpoCod, s.nProSelSubTpo, nNroProceso, c.cConsDescripcion, " & _
            "s.nObjetoCod , s.nMoneda, s.nProSelMonto, s.nProSelEstado, cSintesis, nModalidadCompra " & _
            "from LogProcesoSeleccion s " & _
            "inner join LogProSelTpo t on s.nProSelTpoCod = t.nProSelTpoCod " & _
            "left outer join constante c on s.nObjetoCod=c.nConsValor and c.nConsCod = 9048 " & _
            "where s.nProSelEstado > -1 and s.nNroProceso=" & pnNro & " and nPlanAnualAnio = " & pnAnio
    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sSQL)
        If Not rs.EOF Then
            gnProSelNro = rs!nProselNro
'            gcBSGrupoCod = rs!cBSGrupoCod
            TxtProSelNro.Text = rs!nNroProceso
            TxtTipo.Text = rs!cProSelTpoDescripcion
            TxtMonto.Text = Format(rs!nProSelMonto, "###,###.00")
            LblMoneda.Caption = IIf(rs!nMoneda = 1, "S/.", "$")
            TxtDescripcion.Text = rs!cSintesis
            
            GeneraDetalleItem gnProSelNro
        Else
            FormaFlexItem
            txtPersCod.Text = ""
            txtPersona.Text = ""
            txtMontoGanador.Text = ""
            lblmonedaganador.Caption = ""
            gnProSelNro = 0
            gcBSGrupoCod = ""
            TxtProSelNro.Text = ""
            TxtTipo.Text = ""
            TxtMonto.Text = ""
            LblMoneda.Caption = ""
            TxtDescripcion.Text = ""
            MsgBox "Proceso no Existe", vbInformation
        End If
    End If
    Exit Sub
ConsultarProcesoNroErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(TxtProSelNro.Text) >= 0 Then
            ConsultarProcesoNro Val(TxtProSelNro.Text), Val(txtanio.Text)
            Exit Sub
        Else
            TxtProSelNro.SetFocus
        End If
    End If
    KeyAscii = DigNumEnt(KeyAscii)
End Sub

Private Sub MSItem_DblClick()
On Error GoTo MSItemErr
    Dim i As Integer, bTipo As Boolean
    With MSItem
        If Trim(.TextMatrix(.row, 0)) = "-" Then
           .TextMatrix(.row, 0) = "+"
           i = .row + 1
           bTipo = True
        ElseIf Trim(.TextMatrix(.row, 0)) = "+" Then
           .TextMatrix(.row, 0) = "-"
           i = .row + 1
           bTipo = False
        Else
            Exit Sub
        End If
        
        Do While i < .Rows
            If Trim(.TextMatrix(i, 0)) = "+" Or Trim(.TextMatrix(i, 0)) = "-" Then
                Exit Sub
            End If
            
            If bTipo Then
                .RowHeight(i) = 0
            Else
                .RowHeight(i) = 260
            End If
            i = i + 1
        Loop
    End With
Exit Sub
MSItemErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

'********************************************************************************
' GENERACION DE PROCESOS DE SELECCION
'********************************************************************************

'Sub ListaProcesosMensual(vMes As Integer, vAnio As Integer, vMesDesc As String)
'Dim oConn As New DConecta, Rs As New ADODB.Recordset, i As Integer, nSuma As Currency
'
'sSql = ""
'nSuma = 0
'FormaFlex
''nAnio = CInt(txtAnio.Text)
''nMes = cboMes.ListIndex + 1
'
'If oConn.AbreConexion Then
'
'   'sSQL = "select p.*, t.cProSelTpoDescripcion as cProceso," & _
'          "       d.nPlanAnualAnio,d.cSintesis, d.nPlanAnualMes,d.cBSGrupoCod " & _
'          "  from LogProcesoSeleccion p " & _
'          "       inner join LogProSelTpo t on p.nProSelTpoCod = t.nProSelTpoCod " & _
'          "       inner join LogPlanAnualDetalle d on p.nPlanAnualNro = d.nPlanAnualNro " & _
'          " Where d.nPlanAnualAnio = " & nAnio & " And d.nPlanAnualMes = " & nMes & " and p.nProSelEstado=1"
'
'    sSql = "select p.*, t.cProSelTpoDescripcion as cProceso,i.cBSGrupoCod, " & _
'            "i.cSintesis , P.nPlanAnualMes, i.nProSelItem, i.nMonto " & _
'            "from LogProcesoSeleccion p " & _
'            "inner join LogProSelTpo t on p.nProSelTpoCod = t.nProSelTpoCod " & _
'            "inner join LogProSelItem i on p.nProSelNro = i.nProSelNro " & _
'            "Where p.nPlanAnualAnio = " & nAnio & " And p.nPlanAnualMes = " & nMes & " and p.nProSelEstado=1 "
'
'   If Len(sSql) = 0 Then Exit Sub
'   Set Rs = oConn.CargaRecordSet(sSql)
'   If Not Rs.EOF Then
'      Do While Not Rs.EOF
'         i = i + 1
'         InsRow MSFlex, i
'         MSFlex.RowHeight(i) = 500
'         MSFlex.TextMatrix(i, 0) = Rs!nPlanAnualNro
'         MSFlex.TextMatrix(i, 1) = Rs!nProSelItem
'         MSFlex.TextMatrix(i, 2) = Rs!nProSelTpoCod
'         MSFlex.TextMatrix(i, 3) = Rs!nProSelSubTpo
'         MSFlex.TextMatrix(i, 4) = Rs!nProSelNro
'         MSFlex.TextMatrix(i, 5) = Rs!cSintesis
'         MSFlex.TextMatrix(i, 6) = Rs!cProceso
'         MSFlex.TextMatrix(i, 7) = IIf(Rs!nMoneda = 2, "DOLARES", "SOLES")
'         MSFlex.TextMatrix(i, 8) = FNumero(Rs!nMonto)
'         MSFlex.TextMatrix(i, 9) = Rs!cBSGrupoCod
'         MSFlex.TextMatrix(i, 10) = Rs!cArchivoBases
'         nSuma = nSuma + Rs!nMonto
'         Rs.MoveNext
'      Loop
''      MSflex.SetFocus
'   End If
'End If
'End Sub
'
'Sub FormaFlex()
'MSFlex.Clear
'MSFlex.Rows = 2
'MSFlex.RowHeight(0) = 360
'MSFlex.RowHeight(1) = 8
'MSFlex.ColWidth(0) = 0
'MSFlex.ColWidth(1) = 0:     MSFlex.ColAlignment(1) = 4
'MSFlex.ColWidth(2) = 0
'MSFlex.ColWidth(3) = 0
'MSFlex.ColWidth(4) = 350:   MSFlex.TextMatrix(0, 4) = "Nº":     MSFlex.ColAlignment(4) = 4
'MSFlex.ColWidth(5) = 4000:  MSFlex.TextMatrix(0, 5) = ""
'MSFlex.ColWidth(6) = 3200:  MSFlex.TextMatrix(0, 6) = "Proceso"
'MSFlex.ColWidth(7) = 1000:  MSFlex.TextMatrix(0, 7) = "     Moneda": MSFlex.ColAlignment(7) = 4
'MSFlex.ColWidth(8) = 1100:  MSFlex.TextMatrix(0, 8) = "          Monto"
'MSFlex.WordWrap = True
'End Sub


'********************************************************************************
' GENERACION DE POSTORES GANADORES
'********************************************************************************
'Private Sub MSFlex_GotFocus()
'If Len(Trim(MSFlex.TextMatrix(MSFlex.Row, 4))) > 0 Then
'   GeneraPostoresGanadores Val(MSFlex.TextMatrix(MSFlex.Row, 4)), Val(MSFlex.TextMatrix(MSFlex.Row, 1))
'End If
'End Sub

'Private Sub MSFlex_RowColChange()
'If Len(Trim(MSFlex.TextMatrix(MSFlex.Row, 4))) > 0 Then
'   GeneraPostoresGanadores Val(MSFlex.TextMatrix(MSFlex.Row, 4)), Val(MSFlex.TextMatrix(MSFlex.Row, 1))
'End If
'End Sub

Sub GeneraPostoresGanadores(ByVal pnProSelNro As Integer, ByVal pnProSelItem As Integer)
Dim oConn As New DConecta, rs As New ADODB.Recordset, i As Integer

'FormaFlexPos
'sSQL = "select pp.cPersCod, cPostor=replace(pe.cPersNombre,'/',' '), sum(nPropEconomica) as nMonto" & _
 " from LogProSelPostorPropuesta pp inner join Persona pe on pp.cPersCod = pe.cPersCod " & _
 " Where pp.bGanador = 1 And pp.nProSelNro = " & pnProSelNro & " and pp.nProSelItem= " & pnProSelItem & _
 "  group by pp.cPersCod,pe.cPersNombre "
 sSQL = "select pp.cPersCod, cPostor=replace(pe.cPersNombre,'/',' '), nPropEconomica as nMonto, nProSelConNro=isnull(c.nProSelConNro,0) " & _
        "from LogProSelPostorPropuesta pp " & _
        "    inner join Persona pe on pp.cPersCod = pe.cPersCod " & _
        "    left outer join LogProSelContrato c on pp.nProSelNro = c.nProSelNro and pp.nProSelItem = c.nProSelItem and pp.cPersCod = c.cPersCod " & _
        "    Where pp.bGanador = 1 And pp.nProSelNro = " & pnProSelNro & " and pp.nProSelItem= " & pnProSelItem

If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
Else
   MsgBox "No se puede establecer conexión..." + Space(10), vbInformation
   Exit Sub
End If
If Not rs.EOF Then
'   i = 0
'   Do While Not Rs.EOF
'      i = i + 1
'      InsRow MSPos, i
'      MSPos.TextMatrix(i, 1) = Rs!cPersCod
'      MSPos.TextMatrix(i, 2) = Rs!cPostor
'      MSPos.TextMatrix(i, 3) = FNumero(Rs!nMonto)
'      MSPos.TextMatrix(i, 4) = Rs!nProSelConNro
'      Rs.MoveNext
'   Loop
'   If Not Rs.EOF Then
    txtPersCod.Text = rs!cPersCod
    txtPersona.Text = rs!cPostor
    txtMontoGanador.Text = FNumero(rs!nMonto)
    lblmonedaganador.Caption = LblMoneda.Caption
Else
    txtPersCod.Text = ""
    txtPersona.Text = ""
    txtMontoGanador.Text = ""
    lblmonedaganador.Caption = ""
    MsgBox "No Existe Ganador para el Item del Proceso Seleccionado", vbInformation, "Aviso"
End If
End Sub

'Sub FormaFlexPos()
'MSPos.Clear
'MSPos.Rows = 2
'MSPos.RowHeight(0) = 360
'MSPos.RowHeight(1) = 8
'MSPos.ColWidth(0) = 0
'MSPos.ColWidth(1) = 1050:    MSPos.TextMatrix(0, 1) = "Código":     MSPos.ColAlignment(1) = 4
'MSPos.ColWidth(2) = 3600:    MSPos.TextMatrix(0, 2) = "Postor"
'MSPos.ColWidth(3) = 1100:    MSPos.TextMatrix(0, 3) = " Monto"
'MSPos.ColWidth(4) = 800:     MSPos.TextMatrix(0, 4) = " Contrato"
'End Sub

'********************************************************************************
' GENERACION DE ITEMS POR POSTOR
'********************************************************************************
'Private Sub MSPos_GotFocus()
'If Len(Trim(MSPos.TextMatrix(MSPos.Row, 1))) > 0 Then
'   GeneraItemsPostor Val(MSFlex.TextMatrix(MSFlex.Row, 4)), MSPos.TextMatrix(MSPos.Row, 1), Val(MSFlex.TextMatrix(MSFlex.Row, 1))
'    GeneraItemsPostor Val(MSFlex.TextMatrix(MSFlex.Row, 4)), txtPersCod.Text, Val(MSFlex.TextMatrix(MSFlex.Row, 1))
'End If
'End Sub

'Private Sub MSPos_RowColChange()
'If Len(Trim(MSPos.TextMatrix(MSPos.Row, 1))) > 0 Then
'   GeneraItemsPostor MSFlex.TextMatrix(MSFlex.Row, 4), MSPos.TextMatrix(MSPos.Row, 1), Val(MSFlex.TextMatrix(MSFlex.Row, 1))
'End If
'End Sub

Sub GeneraItemsPostor(ByVal pnProSelNro As Integer, ByVal pcPersCod As String, ByVal pnProSelItem As Integer)
Dim oConn As New DConecta, rs As New ADODB.Recordset, i As Integer

FormaFlexItem

sSQL = "select i.cProSelBSCod, b.cBSDescripcion, i.nCantidad, t.cUnidad,p.cPersCod " & _
        "  from LogProSelItemBS i inner join LogProSelBienesServicios b on i.cProSelBSCod = b.cProSelBSCod " & _
        " inner join (select nConsValor as nBSUnidad, cConsDescripcion as cUnidad from Constante where nConsCod = 9097) t on b.nBSUnidad = t.nBSUnidad " & _
        " inner join (select distinct nProSelNro,nProSelItem,cPersCod from LogProSelPostorPropuesta where bGanador=1) p on p.nProSelNro = i.nProSelNro and p.nProSelItem = i.nProSelItem " & _
        " where i.nProSelNro = " & pnProSelNro & " and p.cPersCod = '" & pcPersCod & "' and i.nProSelItem=" & pnProSelItem

If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
Else
   MsgBox "No se puede establecer conexión..." + Space(10), vbInformation
   Exit Sub
End If

If Not rs.EOF Then
   i = 0
   Do While Not rs.EOF
      i = i + 1
      InsRow MSItem, i
      MSItem.TextMatrix(i, 1) = rs!cProSelBSCod
      MSItem.TextMatrix(i, 2) = rs!cBSDescripcion
      MSItem.TextMatrix(i, 3) = rs!nCantidad
      MSItem.TextMatrix(i, 4) = rs!cUnidad
      rs.MoveNext
   Loop
End If
End Sub

Sub FormaFlexItem()
MSItem.Clear
MSItem.Rows = 2
MSItem.Cols = 10
MSItem.RowHeight(0) = 320
MSItem.RowHeight(1) = 8
MSItem.ColWidth(0) = 250: MSItem.ColAlignment(0) = 4
MSItem.ColWidth(1) = 1000:   MSItem.ColAlignment(1) = 4:  MSItem.TextMatrix(0, 1) = " Item"
MSItem.ColWidth(2) = 0:   MSItem.ColAlignment(2) = 4:  MSItem.TextMatrix(0, 2) = " Código"
MSItem.ColWidth(3) = 8500:  MSItem.TextMatrix(0, 3) = " Descripción"
MSItem.ColWidth(4) = 1000:  MSItem.TextMatrix(0, 4) = "Cant."
MSItem.ColWidth(5) = 0:  MSItem.TextMatrix(0, 5) = " nProSelNro"
MSItem.ColWidth(6) = 0:  MSItem.TextMatrix(0, 6) = " nProSelItem"
MSItem.ColWidth(7) = 0:  MSItem.TextMatrix(0, 7) = "P. Uni."
MSItem.ColWidth(8) = 0:  MSItem.TextMatrix(0, 8) = " Monto"
MSItem.ColWidth(9) = 0:  MSItem.TextMatrix(0, 9) = " Precio"
End Sub

'Private Sub txtanio_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        'cboMes_Click
'    Else
'        KeyAscii = DigNumEnt(KeyAscii)
'    End If
'End Sub

Private Sub MSItem_GotFocus()
    GeneraPostoresGanadores gnProSelNro, Val(MSItem.TextMatrix(MSItem.row, 6))
End Sub

Private Sub MSItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then MSItem_DblClick
End Sub

Private Sub MSItem_SelChange()
    GeneraPostoresGanadores gnProSelNro, Val(MSItem.TextMatrix(MSItem.row, 6))
End Sub
