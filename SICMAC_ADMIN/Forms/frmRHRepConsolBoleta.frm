VERSION 5.00
Begin VB.Form frmRHRepConsolBoleta 
   Caption         =   "Boletas Consolidadas"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10860
   Icon            =   "frmRHRepConsolBoleta.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7980
   ScaleWidth      =   10860
   Begin VB.ComboBox cmbOpc 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8520
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   120
      Width           =   2190
   End
   Begin VB.CommandButton cmdprocesar 
      Caption         =   "Procesar"
      Height          =   375
      Left            =   9240
      TabIndex        =   11
      Top             =   600
      Width           =   1455
   End
   Begin VB.ComboBox cmbmes 
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.ComboBox cmbano 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame fraAgencias 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   8925
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Todos"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   150
         TabIndex        =   5
         Top             =   240
         Width           =   930
      End
      Begin Sicmact.TxtBuscar TxtAgencia 
         Height          =   285
         Left            =   1065
         TabIndex        =   4
         Top             =   210
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
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
         sTitulo         =   ""
      End
      Begin VB.Label lblAgencia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2460
         TabIndex        =   6
         Top             =   195
         Width           =   6375
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9120
      TabIndex        =   2
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   7560
      Width           =   1335
   End
   Begin Sicmact.FlexEdit fgeConsolidaBoleta 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   10398
      Cols0           =   5
      HighLight       =   1
      AllowUserResizing=   1
      EncabezadosNombres=   "Item-Código-Nombres-Agencia-Imprime"
      EncabezadosAnchos=   "450-1500-4000-2500-1200"
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
      ColumnasAEditar =   "X-1-X-X-4"
      TextStyleFixed  =   3
      ListaControles  =   "0-1-0-0-4"
      EncabezadosAlineacion=   "R-L-L-L-C"
      FormatosEdit    =   "0-0-0-0-0"
      CantEntero      =   10
      CantDecimales   =   4
      AvanceCeldas    =   1
      TextArray0      =   "Item"
      lbEditarFlex    =   -1  'True
      lbFlexDuplicados=   0   'False
      lbPuntero       =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   450
      RowHeight0      =   300
   End
   Begin VB.Label lblOpciones 
      Caption         =   "Opc :"
      Height          =   195
      Left            =   7920
      TabIndex        =   13
      Top             =   180
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mes"
      Height          =   195
      Left            =   2520
      TabIndex        =   10
      Top             =   180
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Año"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   180
      Width           =   285
   End
End
Attribute VB_Name = "frmRHRepConsolBoleta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oArea As DActualizaDatosArea
Dim WithEvents oPlaEvento As NActualizaDatosConPlanilla
Attribute oPlaEvento.VB_VarHelpID = -1
Dim Progress As clsProgressBar




Private Sub chkTodos_Click()
    If Me.chkTodos.value = 1 Then
        Me.TxtAgencia.Text = ""
        Me.lblAgencia.Caption = ""
    End If
End Sub



Private Sub cmbOpc_Click()


For i = 1 To fgeConsolidaBoleta.Rows - 1
                fgeConsolidaBoleta.Row = i
                fgeConsolidaBoleta.TextMatrix(i, 4) = Val(Right(cmbOpc.Text, 1))
Next i

End Sub

Private Sub cmdImprimir_Click()
    
    Dim oCon As NConstSistemas
    Set oCon = New NConstSistemas
    Dim PB(8) As Boolean
    Dim lsCadena As String
    Dim lsCadenaExt As String
    Dim i As Long
    Dim sPeriodoAgencia As String
    
   For i = 1 To Me.fgeConsolidaBoleta.Rows - 1
        If fgeConsolidaBoleta.TextMatrix(i, 4) <> "." Then
            lsCadenaExt = lsCadenaExt & "'" & fgeConsolidaBoleta.TextMatrix(i, 1) & "',"
        End If
    Next i
    If lsCadenaExt = "" Then
        lsCadenaExt = "''"
    Else
        lsCadenaExt = lsCadenaExt & "''"
    End If

    sPeriodoAgencia = cmbano.Text + Right(cmbmes.Text, 2) + "----------" + TxtAgencia.Text

    lsCadena = ""
    lsCadena = lsCadena & oPlaEvento.GetBoletas(sPeriodoAgencia, "E01", lsCadenaExt, "PLANILLA CONSOLIDADA ", Date, gsRUC, gsEmpresaCompleto, "PLANILLA CONSOLIDADA " + Left(cmbmes.Text, 8) + " " & cmbano.Text)
  
    
    Dim MSWord As Word.Application
    Dim MSWordSource As Word.Application
    Set MSWord = New Word.Application
    Set MSWordSource = New Word.Application
    Dim RangeSource As Word.Range
    
    MSWordSource.Documents.Open FileName:=App.path & "\SPOOLER\Boletas_Pago.doc"
    Set RangeSource = MSWordSource.ActiveDocument.Content
    'Lo carga en Memoria
    MSWordSource.ActiveDocument.Content.Copy
    'MSWordSource.ActiveDocument
    'Crea Nuevo Documento
    MSWord.Documents.Add
    
    MSWord.Application.Selection.TypeParagraph
    MSWord.Application.Selection.Paste
    MSWord.Application.Selection.InsertBreak
    
    'MSWordSource.ActiveDocument.Close
    Set MSWordSource = Nothing
        
    MSWord.Selection.SetRange start:=MSWord.Selection.start, End:=MSWord.ActiveDocument.Content.End
    MSWord.Selection.MoveEnd
          
                

    MSWord.ActiveDocument.Range.InsertBefore lsCadena
    MSWord.ActiveDocument.Select
    MSWord.ActiveDocument.Range.Font.Name = "Courier New"
    MSWord.ActiveDocument.Range.Font.Size = 6
    MSWord.ActiveDocument.Range.Paragraphs.Space1
    
    MSWord.Selection.Find.Execute Replace:=wdReplaceAll
    MSWord.ActiveDocument.PageSetup.Orientation = wdOrientLandscape
    
    MSWord.ActiveDocument.PageSetup.TopMargin = CentimetersToPoints(2)
    MSWord.ActiveDocument.PageSetup.LeftMargin = CentimetersToPoints(1)
    MSWord.ActiveDocument.PageSetup.RightMargin = CentimetersToPoints(1)

    'MSWord.ActiveDocument.PageSetup.TopMargin = CentimetersToPoints(1)
    'MSWord.ActiveDocument.PageSetup.LeftMargin = CentimetersToPoints(0.5)
    'MSWord.ActiveDocument.PageSetup.RightMargin = CentimetersToPoints(0.5)
    'MSWord.ActiveDocument.PageSetup.TopMargin = CentimetersToPoints(1)
    'MSWord.ActiveDocument.PageSetup.LeftMargin = CentimetersToPoints(0.5)
    'Documento.PageSetup.RightMargin = CentimetersToPoints(0.5)
    
    MSWord.ActiveDocument.SaveAs App.path & "\SPOOLER\Boletas_Pago_" & gsCodUser & Format(Now, "yyyymmsshhmmss") & ".doc"
    MSWord.Visible = True
                
                
 
End Sub

Private Sub cmdProcesar_Click()
Dim sAno As String
Dim sMes As String
Dim sAgencia As String
Dim nimprime As Integer
If cmbano.Text = "" Then Exit Sub
If cmbmes.Text = "" Then Exit Sub

sAno = cmbano.Text
sMes = Right(cmbmes.Text, 2)

sAgencia = TxtAgencia.Text
nimprime = Right(cmbOpc.Text, 1)

Set fgeConsolidaBoleta.Recordset = oPlaEvento.ListadoBoletas(sAno + sMes, sAgencia, nimprime)



End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub


Private Sub Form_Load()
cmbano.AddItem "2004"
cmbano.AddItem "2005"
cmbano.AddItem "2006"

cmbmes.AddItem "ENERO" & Space(20) & "01"
cmbmes.AddItem "FEBRERO" & Space(20) & "02"
cmbmes.AddItem "MARZO" & Space(20) & "03"
cmbmes.AddItem "ABRIL" & Space(20) & "04"
cmbmes.AddItem "MAYO" & Space(20) & "05"
cmbmes.AddItem "JUNIO" & Space(20) & "06"
cmbmes.AddItem "JULIO" & Space(20) & "07"
cmbmes.AddItem "AGOSTO" & Space(20) & "08"
cmbmes.AddItem "SETIEMBRE" & Space(20) & "09"
cmbmes.AddItem "OCTUBRE" & Space(20) & "10"
cmbmes.AddItem "NOVIEMBRE" & Space(20) & "11"
cmbmes.AddItem "DICIEMBRE" & Space(20) & "12"


cmbOpc.AddItem "MARCAR TODOS                                                 1"
cmbOpc.AddItem "DESMARCAR TODOS                                              0"

cmbOpc.ListIndex = 0
cmbmes.ListIndex = 0
cmbano.ListIndex = 1

Set oArea = New DActualizaDatosArea
Me.TxtAgencia.rs = oArea.GetAgencias
Set oPlaEvento = New NActualizaDatosConPlanilla
Set Progress = New clsProgressBar



End Sub


Private Sub txtAgencia_EmiteDatos()
lblAgencia.Caption = TxtAgencia.psDescripcion
End Sub

Private Sub oPlaEvento_CloseProgress()
    Progress.CloseForm Me
End Sub

Private Sub oPlaEvento_Progress(pnValor As Long, pnTotal As Long)
    Progress.Max = pnTotal
    Progress.Progress pnValor, "Generando Reporte"
End Sub

Private Sub oPlaEvento_ShowProgress()
    Progress.ShowForm Me
End Sub

