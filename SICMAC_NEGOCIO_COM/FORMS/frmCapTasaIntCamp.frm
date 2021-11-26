VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCapTasaIntCamp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Captaciones - Tasas Interés Campaña - "
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8580
   Icon            =   "frmCapTasaIntCamp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTarifa 
      Caption         =   " Tasas "
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   4095
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   5
      Top             =   1800
      Width           =   8355
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6960
         TabIndex        =   11
         Top             =   3600
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   10
         Top             =   3600
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1200
         TabIndex        =   9
         Top             =   3600
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2280
         TabIndex        =   8
         Top             =   3600
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6960
         TabIndex        =   7
         Top             =   3600
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5880
         TabIndex        =   6
         Top             =   3600
         Visible         =   0   'False
         Width           =   1035
      End
      Begin SICMACT.FlexEdit grdTasas 
         Height          =   2895
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   5106
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   1
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Monto Ini-Monto Fin-Plazo Ini-Plazo Fin-Ord?-Tasa Int-nTasaCod-Activa-bEdit"
         EncabezadosAnchos=   "300-1500-1500-1200-1200-600-1100-0-0-0"
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-2-3-4-5-6-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-4-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-R-R-R-R-C-R-L-C-C"
         FormatosEdit    =   "0-2-2-3-3-0-2-0-1-1"
         CantEntero      =   12
         CantDecimales   =   4
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label lblAgencias 
         Caption         =   "Ver"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   7800
         MousePointer    =   10  'Up Arrow
         TabIndex        =   19
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Agencias: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6840
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblSubProducto 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   17
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Sub Producto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblMoneda 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Moneda: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7350
      TabIndex        =   4
      Top             =   6075
      Width           =   1155
   End
   Begin VB.Frame fraTasa 
      Caption         =   " Buscar "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   8355
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7200
         TabIndex        =   3
         Top             =   240
         Width           =   1035
      End
      Begin MSComctlLib.ListView LstCampanas 
         Height          =   1455
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nº"
            Object.Width           =   531
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "nCampanaCod"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Campaña"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Moneda"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Sub Producto"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Monto Min"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Fecha Ini"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Fecha Fin"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Campaña:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdNuevaBus 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      Picture         =   "frmCapTasaIntCamp.frx":030A
      TabIndex        =   1
      Top             =   6075
      Width           =   1155
   End
End
Attribute VB_Name = "frmCapTasaIntCamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************************
'** Nombre : frmCapTasaIntCamp
'** Descripción : Formulario para administrar las campañas de tasas de captaciones
'** Creación : JUEZ, 20160415 09:00:00 AM
'*******************************************************************************************

Option Explicit

Dim oNCapDef As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim objPista As COMManejador.Pista
Dim rs As ADODB.Recordset
Dim fsProd As Producto
Dim fsOpeCod As CaptacOperacion
Dim bConsulta As Boolean
Dim nColEdita As Integer
Dim rsTasasAux As ADODB.Recordset

Public Sub inicia(ByVal psProd As Producto, Optional ByVal pbCons As Boolean = False)
    fsProd = psProd
    bConsulta = pbCons
    Select Case fsProd
        Case gCapAhorros
            fsOpeCod = gAhoMantTasaInteresCamp
            Me.Caption = Me.Caption & "Ahorros - " & IIf(bConsulta, "Consulta", "Mantenimiento")
        Case gCapPlazoFijo
            fsOpeCod = gPFMantTasaInteresCamp
            Me.Caption = Me.Caption & "Plazo Fijo - " & IIf(bConsulta, "Consulta", "Mantenimiento")
            grdTasas.ColWidth(5) = 0
        Case gCapCTS
            fsOpeCod = gCTSMantTasaInteresCamp
            Me.Caption = Me.Caption & "CTS - " & IIf(bConsulta, "Consulta", "Mantenimiento")
            grdTasas.ColWidth(5) = 0
    End Select
    CargarCampanas
    Me.Show 1
End Sub

Private Sub CargarCampanas()
Dim L As ListItem
    
    Set oNCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
        Set rs = oNCapDef.GetCaptacCampanas(, fsProd)
    Set oNCapDef = Nothing
    
    LstCampanas.ListItems.Clear
    
    If Not rs.EOF And Not rs.BOF Then
        LstCampanas.Enabled = True
        
        Do While Not rs.EOF
            Set L = LstCampanas.ListItems.Add(, , rs.Bookmark)
            L.SubItems(1) = rs("nCampanaCod")
            L.SubItems(2) = rs("cCampanaDesc")
            L.SubItems(3) = rs("cMoneda")
            L.SubItems(4) = rs("cSubProducto")
            L.SubItems(5) = Format(rs("nMontoMin"), "#,##0.00")
            L.SubItems(6) = Format(rs("dFechaIni"), "dd/MM/yyyy")
            L.SubItems(7) = Format(rs("dFechaFin"), "dd/MM/yyyy")
            rs.MoveNext
        Loop
    Else
        LstCampanas.Enabled = False
        MsgBox "No existen campañas registradas", vbInformation, "Aviso"
        cmdBuscar.Enabled = False
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim i As Integer
    For i = 1 To 6
        If i <> 5 Then
            If grdTasas.TextMatrix(grdTasas.row, i) = "" Then
                MsgBox "Debe ingresar todos los datos", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
    Next i
    
    If Not ValidaTasasNew Then Exit Sub
    
    HabilitaControlesModifGrid (True)
    If cmdAgregar.Enabled And cmdAgregar.Visible Then cmdAgregar.SetFocus
End Sub

Private Sub CmdAgregar_Click()
    HabilitaControlesModifGrid (False)
    
    Set rsTasasAux = grdTasas.GetRsNew
    grdTasas.AdicionaFila
    nColEdita = grdTasas.row
    SendKeys "{Enter}"
End Sub

Private Sub cmdCancelar_Click()
    HabilitaControlesModifGrid (True)
    Set grdTasas.Recordset = rsTasasAux
    If fsProd <> gCapAhorros Then grdTasas.ColWidth(5) = 0
    Set rsTasasAux = Nothing
    If cmdAgregar.Enabled And cmdAgregar.Visible Then cmdAgregar.SetFocus
End Sub

Private Sub cmdGrabar_Click()
Dim i As Integer
Dim nCodTasa As Long
Dim nMontoIni As Double, nMontoFin As Double
Dim nPlazoIni As Long, nPlazoFin As Long
Dim nValorTasa As Double
Dim bOrdPag As Boolean
Dim bActiva As Boolean
Dim sMovNro As String

    If Trim(grdTasas.TextMatrix(1, 1)) = "" Then
        MsgBox "Ingrese las Tasas", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("¿Desea grabar la información actualizada?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    Set oNCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    oNCapDef.ActualizaTasaCampana CLng(LstCampanas.SelectedItem.SubItems(1))
    
    For i = 1 To grdTasas.Rows - 1
        nMontoIni = grdTasas.TextMatrix(i, 1)
        nMontoFin = grdTasas.TextMatrix(i, 2)
        nPlazoIni = grdTasas.TextMatrix(i, 3)
        nPlazoFin = grdTasas.TextMatrix(i, 4)
        nValorTasa = CDbl(grdTasas.TextMatrix(i, 6))
        bOrdPag = IIf(grdTasas.TextMatrix(i, 5) = ".", True, False)
        bActiva = IIf(grdTasas.TextMatrix(i, 8) = ".", True, False)
        sMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        
        oNCapDef.NuevaTasaCampana CLng(LstCampanas.SelectedItem.SubItems(1)), bOrdPag, nPlazoIni, nPlazoFin, nMontoIni, nMontoFin, nValorTasa, sMovNro
    Next i
    Set oNCapDef = Nothing
    
    Set objPista = New COMManejador.Pista
        objPista.InsertarPista fsOpeCod, sMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar
    Set objPista = Nothing
    
    MsgBox "Las tasas fueron actualizadas", vbInformation, "Aviso"
    
    cmdNuevaBus_Click
End Sub

Private Sub CmdModificar_Click()
    If grdTasas.TextMatrix(grdTasas.row, 0) = "" Then
        MsgBox "Debe seleccionar al menos un registro", vbInformation, "Aviso"
        Exit Sub
    End If
    HabilitaControlesModifGrid (False)
    
    Set rsTasasAux = grdTasas.GetRsNew
    nColEdita = grdTasas.row
End Sub

Private Sub cmdQuitar_Click()
    If grdTasas.TextMatrix(grdTasas.row, 0) = "" Then
        MsgBox "Debe seleccionar al menos un registro", vbInformation, "Aviso"
        Exit Sub
    End If
    If MsgBox("¿Está seguro de eliminar los datos de la fila " + CStr(grdTasas.row) + "?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        grdTasas.EliminaFila grdTasas.row
    End If
End Sub

Private Sub cmdBuscar_Click()
    If Not bConsulta Then
        HabilitaControles (True)
    Else
        HabilitaControlesConsulta (True)
    End If
        
    lblMoneda.Caption = LstCampanas.SelectedItem.SubItems(3)
    lblSubProducto.Caption = LstCampanas.SelectedItem.SubItems(4)
    
    Set oNCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
        Set rs = oNCapDef.GetCaptacTasasCampanas(CLng(LstCampanas.SelectedItem.SubItems(1)))
    Set oNCapDef = Nothing
    
    If Not (rs.EOF And rs.BOF) Then
        Set grdTasas.Recordset = rs
        If fsProd <> gCapAhorros Then grdTasas.ColWidth(5) = 0
    End If
    If fraTarifa.Enabled And cmdAgregar.Enabled And cmdAgregar.Visible Then cmdAgregar.SetFocus
End Sub

Private Sub HabilitaControles(ByVal pbHabilita As Boolean)
    fraTasa.Enabled = Not pbHabilita
    fraTarifa.Enabled = pbHabilita
    cmdAgregar.Visible = pbHabilita
    cmdModificar.Visible = pbHabilita
    cmdQuitar.Visible = pbHabilita
    cmdGrabar.Visible = pbHabilita
    cmdAceptar.Visible = Not pbHabilita
    cmdCancelar.Visible = Not pbHabilita
End Sub

Private Sub HabilitaControlesConsulta(ByVal pbHabilita As Boolean)
    fraTasa.Enabled = Not pbHabilita
    fraTarifa.Enabled = pbHabilita
    grdTasas.lbEditarFlex = Not pbHabilita
End Sub

Private Sub HabilitaControlesModifGrid(ByVal pbHabilita As Boolean)
    cmdAgregar.Visible = pbHabilita
    cmdModificar.Visible = pbHabilita
    cmdQuitar.Visible = pbHabilita
    cmdGrabar.Visible = pbHabilita
    cmdAceptar.Visible = Not pbHabilita
    cmdCancelar.Visible = Not pbHabilita
    grdTasas.lbEditarFlex = Not pbHabilita
End Sub

Private Sub cmdNuevaBus_Click()
    If Not bConsulta Then
        HabilitaControles (False)
    Else
        HabilitaControlesConsulta (False)
    End If
    lblMoneda.Caption = ""
    lblSubProducto.Caption = ""
    grdTasas.Clear
    grdTasas.Rows = 2
    grdTasas.FormaCabecera
    If fsProd <> gCapAhorros Then grdTasas.ColWidth(5) = 0
    If LstCampanas.Enabled Then LstCampanas.SetFocus
End Sub

Private Sub fraTarifa_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Screen.MousePointer = 0
End Sub

Private Sub grdTasas_Click()
If grdTasas.row <> nColEdita Then
    grdTasas.lbEditarFlex = False
Else
    grdTasas.lbEditarFlex = True
End If
End Sub

Private Sub grdTasas_DblClick()
If grdTasas.row <> nColEdita Then
    grdTasas.lbEditarFlex = False
Else
    grdTasas.lbEditarFlex = True
End If
End Sub

Private Sub grdTasas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If grdTasas.row <> nColEdita Then
        grdTasas.lbEditarFlex = False
    Else
        grdTasas.lbEditarFlex = True
    End If
End If
End Sub

Private Sub grdTasas_OnCellChange(pnRow As Long, pnCol As Long)
If pnRow <> nColEdita Then
    grdTasas.lbEditarFlex = False
Else
    grdTasas.lbEditarFlex = True
End If

End Sub

Private Sub grdTasas_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
'If grdTasas.TextMatrix(pnRow, 10) = "" Then grdTasas.TextMatrix(pnRow, 10) = "A"
'cmdGrabar.Enabled = True
End Sub

Private Sub grdTasas_OnRowAdd(pnRow As Long)
'grdTasas.TextMatrix(pnRow, 10) = "N"
'cmdGrabar.Enabled = True
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub lblAgencias_Click()
Dim oAge As COMDConstantes.DCOMAgencias
Dim rsLista As ADODB.Recordset, rsDatos As ADODB.Recordset
    
    Screen.MousePointer = 0
    Set oAge = New COMDConstantes.DCOMAgencias
        Set rsLista = oAge.ObtieneAgencias()
    Set oAge = Nothing
    Set oNCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
        Set rsDatos = oNCapDef.GetCaptacCampanasAge(LstCampanas.SelectedItem.SubItems(1))
    Set oNCapDef = Nothing
    frmCredListaDatos.Inicio "Agencias", rsDatos, rsLista, 1
End Sub

Private Function ValidaTasasNew() As Boolean
Dim i As Long, j As Long
Dim nMontoIni As Double, nMontoFin As Double
Dim nPlazoIni As Long, nPlazoFin As Long
Dim nMontoIniAux As Double, nMontoFinAux As Double
Dim nPlazoIniAux As Long, nPlazoFinAux As Long
Dim nOrdPag As Integer, nOrdPagAux As Integer

For i = IIf(grdTasas.Rows - 1 > 1, 2, 1) To grdTasas.Rows - 1
    nMontoIni = CDbl(grdTasas.TextMatrix(i, 1))
    nMontoFin = CDbl(grdTasas.TextMatrix(i, 2))
    If CDbl(grdTasas.TextMatrix(i, 3)) > 999999999 Or CDbl(grdTasas.TextMatrix(i, 4)) > 999999999 Or _
        CDbl(IIf(grdTasas.Rows - 1 > 1, grdTasas.TextMatrix(i - 1, 1), 0)) > 999999999 Or _
        CDbl(IIf(grdTasas.Rows - 1 > 1, grdTasas.TextMatrix(i - 1, 4), 0)) > 999999999 Then
        MsgBox "Los plazos no están correctamente ingresados", vbInformation, "Aviso"
        ValidaTasasNew = False
        Exit Function
    End If
    nPlazoIni = CDbl(grdTasas.TextMatrix(i, 3))
    nPlazoFin = CDbl(grdTasas.TextMatrix(i, 4))
    nOrdPag = IIf(grdTasas.TextMatrix(i, 5) = ".", 1, 0)
    nMontoIniAux = CDbl(IIf(grdTasas.Rows - 1 > 1, grdTasas.TextMatrix(i - 1, 1), 0))
    nMontoFinAux = CDbl(IIf(grdTasas.Rows - 1 > 1, grdTasas.TextMatrix(i - 1, 2), 0))
    nPlazoIniAux = CDbl(IIf(grdTasas.Rows - 1 > 1, grdTasas.TextMatrix(i - 1, 3), 0))
    nPlazoFinAux = CDbl(IIf(grdTasas.Rows - 1 > 1, grdTasas.TextMatrix(i - 1, 4), 0))
    nOrdPagAux = IIf(grdTasas.TextMatrix(i - 1, 5) = ".", 1, 0)
    
    If nMontoFin < nMontoIni Then
        MsgBox "El Monto Final no puede ser menor al Monto Inicial en la fila " & i, vbInformation, "Aviso"
        ValidaTasasNew = False
        Exit Function
    End If
    If nPlazoFin < nPlazoIni Then
        MsgBox "El Plazo Final no puede ser menor al Plazo Inicial en la fila " & i, vbInformation, "Aviso"
        ValidaTasasNew = False
        Exit Function
    End If
Next i

ValidaTasasNew = True
End Function

Private Sub lblAgencias_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Screen.MousePointer = 10
End Sub

