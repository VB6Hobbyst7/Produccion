VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmColBienesAdjudLista 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bienes Adjudicados/Embargados/Vendidos"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13455
   Icon            =   "frmColBienesAdjudLista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   13455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAgreEmb 
      Caption         =   "Agregar Embarg."
      Height          =   375
      Left            =   3000
      TabIndex        =   19
      Top             =   5640
      Width           =   1335
   End
   Begin VB.OptionButton optEmbargados 
      Caption         =   "Bienes embargados"
      Height          =   255
      Left            =   2160
      TabIndex        =   18
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdExpExcel 
      Caption         =   "Exportar"
      Height          =   375
      Left            =   7320
      TabIndex        =   13
      ToolTipText     =   "Permite Exportar el contenido del Grid a Excel."
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "Ver Detalle"
      Height          =   375
      Left            =   12000
      TabIndex        =   5
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdVender 
      Caption         =   "Registrar Venta"
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar Adjud."
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   13215
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   5880
         TabIndex        =   10
         Top             =   120
         Width           =   5535
         Begin VB.CheckBox chkTodos 
            Caption         =   "Mostrar todos"
            Height          =   255
            Left            =   3960
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
         Begin MSMask.MaskEdBox txtFecDel 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   14
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFecAl 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
            Height          =   255
            Left            =   2520
            TabIndex        =   15
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Hasta:"
            Height          =   195
            Left            =   2040
            TabIndex        =   17
            Top             =   240
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Desde:"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   510
         End
      End
      Begin VB.CommandButton cmdListar 
         Caption         =   "Listar"
         Height          =   375
         Left            =   11520
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   5535
         Begin VB.OptionButton optVende 
            Caption         =   "Bienes vendidos"
            Height          =   255
            Left            =   3720
            TabIndex        =   8
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optAdjudica 
            Caption         =   "Bienes adjudicados"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
      End
   End
   Begin SICMACT.FlexEdit FeAdj 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   13215
      _extentx        =   19500
      _extenty        =   5530
      cols0           =   8
      highlight       =   1
      allowuserresizing=   1
      rowsizingmode   =   1
      encabezadosnombres=   "Nº-Agencia-Num.Adju-Descripción-Fec.Adju-Valor Adju.-Capital-Int. y Otros"
      encabezadosanchos=   "400-2500-800-4000-1200-1200-1200-1200"
      font            =   "frmColBienesAdjudLista.frx":030A
      font            =   "frmColBienesAdjudLista.frx":0336
      font            =   "frmColBienesAdjudLista.frx":0362
      font            =   "frmColBienesAdjudLista.frx":038E
      font            =   "frmColBienesAdjudLista.frx":03BA
      fontfixed       =   "frmColBienesAdjudLista.frx":03E6
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1
      columnasaeditar =   "X-X-X-X-X-X-X-X"
      textstylefixed  =   4
      listacontroles  =   "0-0-0-0-0-0-0-0"
      encabezadosalineacion=   "C-C-L-L-R-R-R-R"
      formatosedit    =   "0-0-0-5-5-2-2-2"
      avanceceldas    =   1
      textarray0      =   "Nº"
      selectionmode   =   1
      lbeditarflex    =   -1
      lbformatocol    =   -1
      lbbuscaduplicadotext=   -1
      colwidth0       =   405
      rowheight0      =   300
      forecolorfixed  =   -2147483630
   End
End
Attribute VB_Name = "frmColBienesAdjudLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bListaAdjudicados As Boolean
Dim ApExcel As Variant
Private Sub chkTodos_Click()
    If chkTodos.value = vbUnchecked Then
'        lblAnio.Enabled = True
'        lblMes.Enabled = True
'        txtAnio.Enabled = True
'        cboMes.Enabled = True
        
        Me.txtFecDel.Enabled = True
        Me.txtFecAl.Enabled = True
    Else
'        lblAnio.Enabled = False
'        lblMes.Enabled = False
'        txtAnio.Enabled = False
'        cboMes.Enabled = False
    
        Me.txtFecDel.Enabled = False
        Me.txtFecAl.Enabled = False
        Me.txtFecDel.Text = "__/__/____"
        Me.txtFecAl.Text = "__/__/____"
    End If
End Sub

'JIPR20190520 INICIO
Private Sub cmdAgreEmb_Click()
frmColBienesAdjudicacion.Inicio (4)
End Sub
'JIPR20190520 INICIO

Private Sub cmdAgregar_Click()
    frmColBienesAdjudicacion.Inicio (1)
End Sub

Private Sub cmdDetalle_Click()
    Dim NumAdj As Integer
    If FeAdj.row >= 1 And FeAdj.TextMatrix(FeAdj.row, 2) <> "" Then
        NumAdj = FeAdj.TextMatrix(FeAdj.row, 2)
        If NumAdj >= 0 Then
            'Call frmColBienesAdjudicacion.Inicio(3, NumAdj)
            Call frmColBienesAdjudicacion.Inicio(3, NumAdj, bListaAdjudicados) 'PASI20161103 ERS0572016
        Else
            MsgBox "Debe seleccionar un bien para ver el detalle", vbInformation + vbOKOnly, "SICMACM"
            Exit Sub
        End If
    Else
        MsgBox "No existen datos en el listado", vbInformation + vbOKOnly, "SICMACM"
    End If
End Sub

Private Sub CmdEliminar_Click()
    Dim oCnt As COMNContabilidad.NCOMContFunciones
    Set oCnt = New COMNContabilidad.NCOMContFunciones
    
    Dim NumAdj As Integer
    If FeAdj.row >= 1 And FeAdj.TextMatrix(FeAdj.row, 2) <> "" Then
        NumAdj = FeAdj.TextMatrix(FeAdj.row, 2)
        If NumAdj >= 0 Then
            If MsgBox("Se va a eliminar el registro seleccionado ¿Desea continuar?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
                Call oCnt.CambiaEstadoBienAdjudicado(NumAdj, 0)
                MsgBox "El Registro ha sido eliminado", vbInformation, "Aviso"
                Call cmdListar_Click
            End If
        Else
            MsgBox "Debe seleccionar un bien para proceder", vbInformation + vbOKOnly, "SICMACM"
            Exit Sub
        End If
    Else
        MsgBox "No existen datos en el listado", vbInformation + vbOKOnly, "SICMACM"
    End If
End Sub

Private Sub cmdExpExcel_Click()
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim lnFila As Integer
    
    If FeAdj.row >= 1 And FeAdj.TextMatrix(FeAdj.row, 2) <> "" Then
    
        Set ApExcel = CreateObject("Excel.application")
        '-------------------------------
        'Agrega un nuevo Libro
        ApExcel.Workbooks.Add
        'Poner Titulos
        ApExcel.Cells(1, 1) = "CAJA MAYNAS S.A."
        'ApExcel.Cells(2, 1) = IIf(Me.optAdjudica.value = True, "BIENES ADJUDICADOS", "BIENES VENDIDOS") JIPR20190520 COMENTÓ
        ApExcel.Cells(2, 1) = IIf(Me.optAdjudica.value = True, "BIENES ADJUDICADOS", IIf(Me.optVende.value = True, "BIENES VENDIDOS", "BIENES EMBARGADOS")) 'JIPR20190520 AGREGÓ
        ApExcel.Cells(4, 1) = "AGENCIA"
        ApExcel.Cells(4, 2) = "NUM."
        ApExcel.Cells(4, 3) = "DESCRIPCION"
        ApExcel.Cells(4, 4) = "FECHA"
        ApExcel.Cells(4, 5) = "VALOR"
        ApExcel.Cells(4, 6) = "CAPITAL"
        ApExcel.Cells(4, 7) = "INT. Y OTROS"

        ApExcel.Range("A1:G4").Font.Bold = True
        ApExcel.Range("A4:G4").Interior.ColorIndex = 42

        lnFila = 5
        For i = 1 To FeAdj.rows - 1
            ApExcel.Cells(lnFila, 1) = FeAdj.TextMatrix(i, 1)
            ApExcel.Cells(lnFila, 2) = FeAdj.TextMatrix(i, 2)
            ApExcel.Cells(lnFila, 3) = IIf(Asc(Right(FeAdj.TextMatrix(i, 3), 1)) = 10, Left(FeAdj.TextMatrix(i, 3), Len(FeAdj.TextMatrix(i, 3)) - 2), Trim(FeAdj.TextMatrix(i, 3)))
            ApExcel.Cells(lnFila, 4) = "'" + FeAdj.TextMatrix(i, 4)
            ApExcel.Cells(lnFila, 5).NumberFormat = "#,##0.00"
            ApExcel.Cells(lnFila, 5) = FeAdj.TextMatrix(i, 5)
            ApExcel.Cells(lnFila, 6).NumberFormat = "#,##0.00"
            ApExcel.Cells(lnFila, 6) = FeAdj.TextMatrix(i, 6)
            ApExcel.Cells(lnFila, 7).NumberFormat = "#,##0.00"
            ApExcel.Cells(lnFila, 7) = FeAdj.TextMatrix(i, 7)
            lnFila = lnFila + 1
        Next i

        ApExcel.Cells.Select
        ApExcel.Cells.EntireColumn.AutoFit

        ApExcel.Cells.Select
        ApExcel.Cells.Font.Size = 8
        ApExcel.Cells.Range("A1").Select

        ApExcel.Cells.Columns("A:A").ColumnWidth = 20
        ApExcel.Cells.Columns("B:B").ColumnWidth = 12
        ApExcel.Cells.Columns("C:C").ColumnWidth = 70
        ApExcel.Cells.Columns("D:D").ColumnWidth = 13
                
        '-------------------------------
        ApExcel.Visible = True
        Set ApExcel = Nothing
    
    Else
        MsgBox "No existen datos para mostrar.", vbInformation + vbOKOnly, "SICMACM"
    End If

End Sub

Private Sub cmdListar_Click()
'    If Me.optAdjudica.value = True Then
'        bListaAdjudicados = True
'    Else
'        bListaAdjudicados = False
'    End If
'JIPR20190520 COMENTÓ
    
'JIPR20190520 INICIO
    If Me.optAdjudica.value = True Then
        Me.optAdjudica.value = True
        Me.optVende.value = False
        Me.optEmbargados.value = False
    ElseIf Me.optVende.value = True Then
        Me.optAdjudica.value = False
        Me.optVende.value = True
        Me.optEmbargados.value = False
    Else
        Me.optAdjudica.value = False
        Me.optVende.value = False
        Me.optEmbargados.value = True
    End If
'JIPR20190520 FIN
    
'    If chkTodos.value = Unchecked And txtAnio.Text = "" Then
'        MsgBox "Debe ingresar el año del período", vbInformation, "Aviso"
'        Exit Sub
'    End If
'*** PEAC 20130617
    If chkTodos.value = Unchecked Then
        If (IsDate(Me.txtFecDel.Text) = False Or IsDate(Me.txtFecAl.Text) = False) Then
            MsgBox "Ingrese un rango de fechas correctas.", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    If chkTodos.value = Unchecked Then
        If CDate(Me.txtFecDel.Text) > CDate(Me.txtFecAl.Text) Then
            MsgBox "La fecha inicial no puede ser mayor a la fecha final.", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    
    CargaDatos
End Sub

Private Sub cmdVender_Click()
    Dim NumAdj As Integer

    If FeAdj.row >= 1 And FeAdj.TextMatrix(FeAdj.row, 2) <> "" Then
        NumAdj = FeAdj.TextMatrix(FeAdj.row, 2)
        If NumAdj >= 0 Then
            Call frmColBienesAdjudicacion.Inicio(2, NumAdj)
        Else
            MsgBox "Debe seleccionar un bien adjudicado para registrar venta", vbInformation + vbOKOnly, "SICMACM"
            Exit Sub
        End If
    Else
        MsgBox "No existen datos en el listado", vbInformation + vbOKOnly, "SICMACM"
    End If
End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim oGen As DGeneral
    Set oGen = New DGeneral

'*** PEAC 20130618
'    Set rs = oGen.GetConstante(1010)
''    Me.cboMes.Clear
'    While Not rs.EOF
'        cboMes.AddItem rs.Fields(0) & Space(50) & rs.Fields(1)
'        rs.MoveNext
'    Wend
'    Me.txtAnio.Text = Format(gdFecSis, "yyyy")
'    Me.cboMes.ListIndex = Month(gdFecSis) - 1
'*** FIN PEAC

    Me.cmdVender.Visible = True
    If gsCodArea <> "042" Then '042 = Recuperaciones
        Me.cmdAgregar.Visible = False
        Me.cmdCancelar.Visible = False
        Me.cmdVender.Visible = False
        Me.cmdEliminar.Visible = False
        Me.cmdAgreEmb.Visible = False 'JIPR20190520 AGREGÓ
    End If
End Sub

Private Sub CargaDatos()
    Dim oCnt As COMNContabilidad.NCOMContFunciones
    Dim rs As ADODB.Recordset
    Dim sAnio, sMes As String
    Dim tipo As Integer
    Set oCnt = New COMNContabilidad.NCOMContFunciones
    
'    If Me.chkTodos.value = vbChecked Then
'        sAnio = "":   sMes = ""
'    Else
'        sMes = Format(Right(Trim(Me.cboMes.Text), 2), "00")
'        sAnio = Me.txtAnio.Text
'    End If
    
    '*** PEAC 20130617
    'Set rs = oCnt.ObtenerListaBienesAdjudicado(sMes, sAnio, IIf(Me.optAdjudica.value = True, 1, IIf(Me.optAdjudica.value = True, 1, 2)))
    'JIPR20190520 COMENTÓ
   '*** FIN PEAC
    Set rs = oCnt.ObtenerListaBienesAdjudicado(IIf(Me.chkTodos.value = vbChecked, "", Format(Me.txtFecDel.Text, "yyyymmdd")), IIf(Me.chkTodos.value = vbChecked, "", Format(Me.txtFecAl.Text, "yyyymmdd")), IIf(Me.optAdjudica.value = True, 1, IIf(Me.optVende.value = True, 2, 3)))
    If rs.EOF Then MsgBox "No se encontraron datos.", vbInformation, "Mensaje"
    FeAdj.Clear
    FeAdj.FormaCabecera
    FeAdj.rows = 2
    FeAdj.rsFlex = rs
    Set oCnt = Nothing
End Sub

Private Sub optAdjudica_Click()
    If optAdjudica.value = True And gsCodArea = "042" Then
        Me.cmdVender.Visible = True
        
        'JIPR20190520 INICIO
        Me.optVende.value = False
        Me.optEmbargados.value = False
        'JIPR20190520 FIN
                        
    End If
    FeAdj.Clear
    FeAdj.FormaCabecera
    FeAdj.rows = 2
End Sub

'JIPR20190520 AGREGÓ
Private Sub optEmbargados_Click()
 If optEmbargados.value = True And gsCodArea = "042" Then
        Me.cmdVender.Visible = True
        
        'JIPR20190520 INICIO
        Me.optVende.value = False
        Me.optAdjudica.value = False
        'JIPR20190520 FIN
        
    End If
    FeAdj.Clear
    FeAdj.FormaCabecera
    FeAdj.rows = 2
End Sub
'JIPR20190520 AGREGÓ

Private Sub optVende_Click()
    If optVende.value = True And gsCodArea = "042" Then
        Me.cmdVender.Visible = False
        
        'JIPR20190520 INICIO
        Me.optEmbargados.value = False
        Me.optAdjudica.value = False
        'JIPR20190520 FIN
        
    End If
    FeAdj.Clear
    FeAdj.FormaCabecera
    FeAdj.rows = 2
End Sub

Private Sub txtFecDel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtFecAl.SetFocus
    End If
End Sub
