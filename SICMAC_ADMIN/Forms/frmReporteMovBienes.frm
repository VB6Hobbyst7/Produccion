VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmReporteMovBienes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte Movimientos Bienes"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10245
   Icon            =   "frmReporteMovBienes.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   30
      TabIndex        =   17
      Top             =   1320
      Width           =   10095
      Begin MSComctlLib.ListView lvwMovBienes 
         Height          =   3480
         Left            =   30
         TabIndex        =   18
         Top             =   120
         Width           =   10020
         _ExtentX        =   17674
         _ExtentY        =   6138
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   15
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Año"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Codigo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Serie"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Descripcion"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Agencia Origen"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Area Origen"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Agencia Destino"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Area Destino"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Valor Reg"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Valor Ini Cnt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Valor Depre Cnt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Valor por Depre Cnt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Num Movimiento"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   7560
      TabIndex        =   16
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "Extornar (solo estadístico)"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8880
      TabIndex        =   14
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Frame fraCriterio 
      Caption         =   "Criterio"
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
      Height          =   1290
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   10065
      Begin VB.ComboBox cboTpo 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   240
         Width           =   5340
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo de Operación"
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   2655
         Begin VB.OptionButton optBajas 
            Caption         =   "Bajas"
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optTransf 
            Caption         =   "Transferencias"
            Height          =   315
            Left            =   1080
            TabIndex        =   12
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.CheckBox chkGeneral 
         Caption         =   "Todos"
         Height          =   195
         Left            =   6360
         TabIndex        =   2
         Top             =   330
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   7830
         TabIndex        =   1
         Top             =   705
         Width           =   1215
      End
      Begin MSMask.MaskEdBox mskFechaF 
         Height          =   330
         Left            =   6285
         TabIndex        =   4
         Top             =   750
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFechaI 
         Height          =   315
         Left            =   3960
         TabIndex        =   5
         Top             =   750
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2820
         TabIndex        =   7
         Top             =   795
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5280
         TabIndex        =   6
         Top             =   795
         Width           =   1005
      End
      Begin VB.Label lblBien 
         AutoSize        =   -1  'True
         Caption         =   "Bien:"
         Height          =   195
         Left            =   300
         TabIndex        =   3
         Top             =   270
         Width           =   360
      End
   End
   Begin Sicmact.TxtBuscar txtAgeO 
      Height          =   345
      Left            =   1260
      TabIndex        =   8
      Top             =   6960
      Width           =   1455
      _extentx        =   2566
      _extenty        =   609
      appearance      =   0
      appearance      =   0
      font            =   "frmReporteMovBienes.frx":030A
      enabled         =   0
      appearance      =   0
      enabledtext     =   0
   End
   Begin VB.Label lblAgeOG 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   2820
      TabIndex        =   10
      Top             =   6990
      Width           =   4710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Area/Agencias"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   7020
      Width           =   1065
   End
End
Attribute VB_Name = "frmReporteMovBienes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
'*******************************

Private Sub chkGeneral_Click()
If chkGeneral.value = 1 Then
      
   'txtBS.Text = ""
'   lblBienG.Caption = ""
'   txtSerie.Text = ""
   
   txtAgeO.Enabled = True
'   txtBS.Enabled = False
'   txtSerie.Enabled = False
Else
    txtAgeO.Enabled = False
'    txtBS.Enabled = True
'    txtSerie.Enabled = True
    
    txtAgeO.Text = ""
    lblAgeOG.Caption = ""
        
End If

End Sub

Private Sub cmdBuscar_Click()

Dim oALmacen As DLogAlmacen
'Dim rs As ADODB.Recordset
Dim lista As ListItem
Dim parametro As Integer
Dim lsAreaCod As String
Dim lsAgeCod As String
Dim FecIni As String
Dim FecFin As String
Dim J As Integer
Set rs = New ADODB.Recordset
Set oALmacen = New DLogAlmacen

Dim lcOpeTpo As String


'lista.Index()

 If ValidarFechas = False Then Exit Sub
    
    If chkGeneral.value = 1 Then
'        If txtAgeO.Text = "" Then
'           MsgBox "Debe Seleccionar Area", vbInformation, "Aviso"
'           txtAgeO.SetFocus
'           Exit Sub
'        End If
        parametro = 1
    Else
        'If txtBS.Text = "" Then
        If cboTpo.Text = "" Then
        
           MsgBox "Debe Seleccionar un Bien", vbInformation, "Aviso"
           cboTpo.SetFocus
           'txtBS.SetFocus
           Exit Sub
        End If
        parametro = 0
    End If

    If Len(Trim(Me.txtAgeO.Text)) = 3 Then
        lsAreaCod = Trim(txtAgeO.Text)
        lsAgeCod = ""
    Else
       ' lsAreaCod = Mid(Me.txtAgeO.Text, 4, 2)
        lsAgeCod = Mid(Me.txtAgeO.Text, 4, 2)
        lsAreaCod = Left(Me.txtAgeO.Text, 3)
    End If
FecIni = Replace(CStr(Format(mskFechaI, "yyyymmdd")), "/", "")
FecFin = Replace(CStr(Format(mskFechaF, "yyyymmdd")), "/", "")

lcOpeTpo = IIf(Me.optBajas.value = True, "581299", _
            IIf(Me.optTransf.value = True, "581204", ""))

    If lcOpeTpo = "" Then
        MsgBox "Debe Seleccionar un Tipo de Operación.", vbInformation, "Aviso"
        Exit Sub
    End If

    'Set rs = oALmacen.GetLogMovBienes(FecIni, FecFin, parametro, lsAreaCod, lsAgeCod, txtBS.Text, txtSerie.Text, lcOpeTpo) 'Bien x Bien
    Set rs = oALmacen.GetLogMovBienes(FecIni, FecFin, parametro, lsAreaCod, lsAgeCod, , , lcOpeTpo, Right(Me.cboTpo.Text, 3))   'Bien x Bien
    'Set rs = oALmacen.GetLogMovBienes(Mid(mskFechaI, 7, 4), Mid(mskFechaF, 7, 4), parametro, lsAreaCod, lsAgeCod, txtBS.Text, txtSerie.Text) ' Bien x Bien
    Set oALmacen = Nothing

    lvwMovBienes.ListItems.Clear

    '*** PEAC 20120504
    If lcOpeTpo = "581299" Then 'BAJAS
        lvwMovBienes.ColumnHeaders.item(7).Width = 0
        lvwMovBienes.ColumnHeaders.item(8).Width = 0
        lvwMovBienes.ColumnHeaders.item(9).Width = 0
        lvwMovBienes.ColumnHeaders.item(10).Width = 0
        
        lvwMovBienes.ColumnHeaders.item(13).Text = "Valor Depre Cnt"
        lvwMovBienes.ColumnHeaders.item(14).Text = "Valor por Depre Cnt"
        
        'me.FeAdj.
        
    ElseIf lcOpeTpo = "581204" Then 'TRANSFE
        lvwMovBienes.ColumnHeaders.item(7).Width = 1399.74
        lvwMovBienes.ColumnHeaders.item(8).Width = 1399.74
        lvwMovBienes.ColumnHeaders.item(9).Width = 1399.74
        lvwMovBienes.ColumnHeaders.item(10).Width = 1399.74
        
        lvwMovBienes.ColumnHeaders.item(13).Text = "Depre Anteri"
        lvwMovBienes.ColumnHeaders.item(14).Text = "Depre Actual"
    End If

    '*** FIN PEAC
    
    If Not (rs.EOF And rs.BOF) Then
    J = 1
        Do Until rs.EOF
            Set lista = lvwMovBienes.ListItems.Add(, , J)
            
            lista.SubItems(1) = rs(0)
            lista.SubItems(2) = rs(1)
            lista.SubItems(3) = rs(2)
            lista.SubItems(4) = rs(3)
            lista.SubItems(5) = rs(4)
            lista.SubItems(6) = rs(5)
            lista.SubItems(7) = rs(6)
            lista.SubItems(8) = IIf(IsNull(rs(7)), "", rs(7))
            lista.SubItems(9) = rs(8)
            '*** PEAC 20120510
            lista.SubItems(10) = rs(9)
            lista.SubItems(11) = rs(10)
            lista.SubItems(12) = rs(11)
            lista.SubItems(13) = rs(12)
            lista.SubItems(14) = rs(13)
            '*** FIN PEAC
            J = J + 1
            rs.MoveNext
        Loop
    Else
        
        MsgBox "No existen Datos", vbInformation, "Aviso"
    End If
End Sub

Private Sub CmdExtornar_Click()

Dim oMov As DMov
Set oMov = New DMov

Dim lnV1, lnV2, lnV3 As Double
    '581299 bajas - 581204 transf
    lnV1 = Me.lvwMovBienes.SelectedItem.ListSubItems.item(11): lnV2 = Me.lvwMovBienes.SelectedItem.ListSubItems.item(12): lnV3 = Me.lvwMovBienes.SelectedItem.ListSubItems.item(13)
    'MsgBox Me.lvwMovBienes.SelectedItem.ListSubItems.Item(14)

    If CDbl(lnV1) + CDbl(lnV2) + CDbl(lnV3) > 0 Then
        MsgBox "Este registro tiene que extornarlo contablemente en el Financiero, coordine con el Area de Contabilidad.", vbInformation + vbOKOnly, "Atención"
        Exit Sub
    End If

    If MsgBox("¿Está seguro de extornar este movimeinto?", vbQuestion + vbYesNo, "Pregunta") = vbNo Then Exit Sub

    oMov.ModificaActivosFijos Me.lvwMovBienes.SelectedItem.ListSubItems.item(14), IIf(Me.optBajas.value = True, "581299", "581204")

    cmdBuscar_Click
        
        'ARLO 20160126 ***
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", "Se Extorno  el Movimiento"
        Set objPista = Nothing
        '**************

End Sub

Private Sub CmdImprimir_Click()

Dim fs              As Scripting.FileSystemObject
Dim xlAplicacion    As Excel.Application
Dim xlLibro         As Excel.Workbook
Dim xlHoja1         As Excel.Worksheet
Dim lbExisteHoja    As Boolean
Dim liLineas        As Integer
Dim i               As Integer
Dim glsArchivo      As String
Dim lsNomHoja       As String

    
If lvwMovBienes.ListItems.Count < 1 Then
    MsgBox "No exite informacion para imprimir", vbInformation, "Aviso"
    Exit Sub
End If
    
    glsArchivo = "ReporteMovimientosActivosFijos" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
    Set fs = New Scripting.FileSystemObject

    Set xlAplicacion = New Excel.Application
    If fs.FileExists(App.path & "\SPOOLER\" & glsArchivo) Then
        Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & glsArchivo)
    Else
        Set xlLibro = xlAplicacion.Workbooks.Add
    End If
    Set xlHoja1 = xlLibro.Worksheets.Add

    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 60
    xlHoja1.PageSetup.Orientation = xlLandscape

            lbExisteHoja = False
            lsNomHoja = "Adeudos Vinculados"
            For Each xlHoja1 In xlLibro.Worksheets
                If xlHoja1.Name = lsNomHoja Then
                    xlHoja1.Activate
                    lbExisteHoja = True
                    Exit For
                End If
            Next
            If lbExisteHoja = False Then
                Set xlHoja1 = xlLibro.Worksheets.Add
                xlHoja1.Name = lsNomHoja
            End If

            xlAplicacion.Range("A1:A1").ColumnWidth = 5
            xlAplicacion.Range("B1:B1").ColumnWidth = 15
            xlAplicacion.Range("c1:c1").ColumnWidth = 8
            xlAplicacion.Range("D1:D1").ColumnWidth = 10
            xlAplicacion.Range("E1:E1").ColumnWidth = 20
            xlAplicacion.Range("F1:F1").ColumnWidth = 30
            xlAplicacion.Range("G1:G1").ColumnWidth = 20
            xlAplicacion.Range("H1:H1").ColumnWidth = 20
            xlAplicacion.Range("I1:I1").ColumnWidth = 20
            xlAplicacion.Range("J1:J1").ColumnWidth = 20
            
            
            xlAplicacion.Range("K1:K1").ColumnWidth = 20
            xlAplicacion.Range("L1:L1").ColumnWidth = 20
            xlAplicacion.Range("M1:M1").ColumnWidth = 20
            xlAplicacion.Range("N1:N1").ColumnWidth = 20
            xlAplicacion.Range("O1:O1").ColumnWidth = 20
           
            xlAplicacion.Range("A1:Z100").Font.Size = 9
       
            xlHoja1.Cells(1, 1) = gsNomCmac
            xlHoja1.Cells(2, 1) = "Reporte Movimientos de Activos Fijos " & Format(gdFecSis, "dd/mm/yyyy")
            
            xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 3)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 3)).Merge True
            xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 2)).Merge True
                                       
            liLineas = 4
            
            xlHoja1.Cells(liLineas, 1) = "N°"
            xlHoja1.Cells(liLineas, 2) = "Fecha"
            xlHoja1.Cells(liLineas, 3) = "Año"
            xlHoja1.Cells(liLineas, 4) = "Codigo"
            xlHoja1.Cells(liLineas, 5) = "serie"
            xlHoja1.Cells(liLineas, 6) = "Descripcion"
            xlHoja1.Cells(liLineas, 7) = "Agencia Origen"
            xlHoja1.Cells(liLineas, 8) = "Area Origen"
            xlHoja1.Cells(liLineas, 9) = "Agencia Destino"
            xlHoja1.Cells(liLineas, 10) = "Area Destino"
            xlHoja1.Cells(liLineas, 11) = "Valor Registro"
            xlHoja1.Cells(liLineas, 12) = "Valor Ini Cnt"
            xlHoja1.Cells(liLineas, 13) = "Valor Depre Dnt"
            xlHoja1.Cells(liLineas, 14) = "Valor Por Depre Cnt"
            xlHoja1.Cells(liLineas, 15) = "Movi.Cont."

            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas + 3, 1)).Font.Bold = True
            
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 15)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 15)).VerticalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas + 3, 1)).Merge True
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 15)).EntireRow.AutoFit
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 15)).WrapText = True
            
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 15)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 15)).Borders.LineStyle = 1
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 15)).Interior.Color = RGB(159, 206, 238)
            
         
            
            liLineas = liLineas + 1
         rs.MoveFirst
         Dim J As Integer
         J = 1
         Do Until rs.EOF
            xlHoja1.Cells(liLineas, 1) = J
            xlHoja1.Cells(liLineas, 2) = rs(0)
            xlHoja1.Cells(liLineas, 3) = rs(1)
            xlHoja1.Cells(liLineas, 4) = rs(2)
            xlHoja1.Cells(liLineas, 5) = rs(3)
            xlHoja1.Cells(liLineas, 6) = rs(4)
            xlHoja1.Cells(liLineas, 7) = rs(5)
            xlHoja1.Cells(liLineas, 8) = rs(6)
            xlHoja1.Cells(liLineas, 9) = rs(7)
            xlHoja1.Cells(liLineas, 10) = rs(8)
            
            xlHoja1.Cells(liLineas, 11) = rs(9)
            xlHoja1.Cells(liLineas, 12) = rs(10)
            xlHoja1.Cells(liLineas, 13) = rs(11)
            xlHoja1.Cells(liLineas, 14) = rs(12)
            xlHoja1.Cells(liLineas, 15) = rs(13)
            
            J = J + 1
            'xlHoja1.Range(xlHoja1.Cells(liLineas, 3), xlHoja1.Cells(liLineas, 5)).Style = "Comma"
            xlHoja1.Range(xlHoja1.Cells(liLineas, 11), xlHoja1.Cells(liLineas, 14)).Style = "Comma"
            
'            xlHoja1.Range(xlHoja1.Cells(liLineas, 3), xlHoja1.Cells(liLineas, 5)).HorizontalAlignment = xlCenter
'            xlHoja1.Range(xlHoja1.Cells(liLineas, 7), xlHoja1.Cells(liLineas, 7)).HorizontalAlignment = xlCenter
'            xlHoja1.Range(xlHoja1.Cells(liLineas, 10), xlHoja1.Cells(liLineas, 11)).HorizontalAlignment = xlCenter
            
            xlHoja1.Range(xlHoja1.Cells(liLineas, 6), xlHoja1.Cells(liLineas, 6)).HorizontalAlignment = xlRight
            xlHoja1.Range(xlHoja1.Cells(liLineas, 10), xlHoja1.Cells(liLineas, 10)).HorizontalAlignment = xlRight
            
            liLineas = liLineas + 1
            rs.MoveNext
        Loop

       'ExcelCuadro xlHoja1, 1, 4, 12, liLineas - 1
        
        xlHoja1.SaveAs App.path & "\SPOOLER\" & glsArchivo
        ExcelEnd App.path & "\Spooler\" & glsArchivo, xlAplicacion, xlLibro, xlHoja1
    
        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing
        MsgBox "Se ha generado el Archivo en " & App.path & "\SPOOLER\" & glsArchivo
        Call CargaArchivo(glsArchivo, App.path & "\SPOOLER\")
        
        'ARLO 20160126 ***
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Imprimio el Reporte de Movimiento de Bienes "
        Set objPista = Nothing
        '**************
  
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Command1_Click()
    MsgBox Me.lvwMovBienes.SelectedItem.ListSubItems.item(4)
    
    
End Sub

Private Sub Form_Load()
Dim oArea As DActualizaDatosArea
Set oArea = New DActualizaDatosArea
Dim oALmacen As DLogAlmacen
Set oALmacen = New DLogAlmacen

Dim oGen As DGeneral
Set oGen = New DGeneral

Me.txtAgeO.rs = oArea.GetAgenciasAreas
'Me.txtBS.rs = oALmacen.GetAFBienes
'Me.txtBS.rs = oALmacen.GetBienesAlmacen

Set rs = oGen.GetConstante(5062, False)
Me.cboTpo.Clear
While Not rs.EOF
    cboTpo.AddItem rs.Fields(0) & Space(50) & rs.Fields(1)
    rs.MoveNext
Wend


mskFechaI.Text = gdFecSis
mskFechaF.Text = gdFecSis
End Sub

Private Sub mskFechaF_GotFocus()
    mskFechaF.SelStart = 0
    mskFechaF.SelLength = 50
End Sub

Private Sub mskFechaF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdBuscar.SetFocus
    End If
End Sub

Private Sub mskFechaI_GotFocus()
    mskFechaI.SelStart = 0
    mskFechaI.SelLength = 50
End Sub

Private Sub mskFechaI_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        mskFechaF.SetFocus
  End If
End Sub

Private Sub txtAgeO_EmiteDatos()
Me.lblAgeOG.Caption = txtAgeO.psDescripcion
End Sub

Public Function ValidarFechas() As Boolean
     ValidarFechas = True
     If Not IsDate(mskFechaI) = True Then
        MsgBox "Ingrese Fecha Correcta", vbInformation, "Aviso"
        ValidarFechas = False
        Exit Function
     End If
     
     If Not IsDate(mskFechaF) = True Then
        MsgBox "Ingrese Fecha Correcta", vbInformation, "Aviso"
        ValidarFechas = False
        Exit Function
     End If
     If CDate(mskFechaI) > CDate(mskFechaF) Then
        MsgBox "La Fecha incial debe de ser Menor", vbInformation, "Aviso"
        ValidarFechas = False
        Exit Function
     End If
     
End Function

'Private Sub txtBS_EmiteDatos()
'Dim oALmacen As DLogAlmacen
'    Set oALmacen = New DLogAlmacen
'
'    lvwMovBienes.ListItems.Clear
'
'    If txtBS.Text <> "" Then
'        lblBienG.Caption = txtBS.psDescripcion
'        txtSerie.Text = ""
'        Me.txtSerie.rs = oALmacen.GetAFBSSerie(txtBS.Text, Year(gdFecSis))
'    End If
'
'    Set oALmacen = Nothing
'End Sub



