VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPersLavDinero 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   Icon            =   "frmPersLavDinero.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraHistoria 
      Caption         =   "Historia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2535
      Left            =   60
      TabIndex        =   2
      Top             =   3480
      Width           =   9915
      Begin VB.CommandButton cmdComentario 
         Caption         =   "&Editar Com."
         Height          =   375
         Left            =   6960
         TabIndex        =   4
         Top             =   2085
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdNuevoEst 
         Caption         =   "&Nuevo Est."
         Height          =   375
         Left            =   8085
         TabIndex        =   3
         Top             =   2085
         Visible         =   0   'False
         Width           =   1035
      End
      Begin SICMACT.FlexEdit grdHistoria 
         Height          =   2145
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   9600
         _ExtentX        =   16933
         _ExtentY        =   3784
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Fecha-Estado-Comentario"
         EncabezadosAnchos=   "350-1200-2000-5500"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0"
         EncabezadosAlineacion=   "C-C-L-L"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         Appearance      =   0
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame fraPersona 
      Caption         =   "Personas Registradas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3495
      Left            =   60
      TabIndex        =   0
      Top             =   -15
      Width           =   9915
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   315
         Left            =   8640
         TabIndex        =   13
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txtComentario 
         Height          =   285
         Left            =   4080
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   3135
      End
      Begin MSDataListLib.DataCombo dcEstado 
         Height          =   315
         Left            =   4080
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.CommandButton cmdexaminar 
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   7
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Enabled         =   0   'False
         Height          =   315
         Left            =   3720
         TabIndex        =   6
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmdAgregarPers 
         Caption         =   "Agregar "
         Height          =   315
         Left            =   4920
         TabIndex        =   1
         Top             =   3000
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPersona 
         Height          =   2655
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   4683
         _Version        =   393216
         Cols            =   6
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
End
Attribute VB_Name = "frmPersLavDinero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnColumna As Integer
Dim lnFila As Integer
Dim nModificar As Integer '0=inactivo;1=modificar;2=nuevo
Dim nBuscar As Integer '0=inactivo;1=activo
Dim rsGrillaTemp As Recordset
Dim lsEstado As String
Dim lsComentario As String

Private Sub ClearScreen()
grdPersona.Clear
grdPersona.Rows = 2
'grdPersona.FormaCabecera
grdHistoria.Clear
grdHistoria.Rows = 2
grdHistoria.FormaCabecera
Me.cmdGrabar.Enabled = False
Me.cmdCancelar.Enabled = False
Me.cmdexaminar.Enabled = True
lnFila = 0
lnColumna = 0
nModificar = 0
nBuscar = 0
ObtienePersonas
Me.cmdexaminar.SetFocus
End Sub

Private Function ValidaDatosGrid() As Boolean
    ValidaDatosGrid = True
    
    If dcEstado.Visible = True Then
        MsgBox "Presione Enter para seleccionar el Estado", vbInformation, "Aviso"
        ValidaDatosGrid = False
        dcEstado.SetFocus
        Exit Function
    End If
    If txtComentario.Visible = True Then
        MsgBox "Presione Enter para ingresar el Comentario", vbInformation, "Aviso"
        ValidaDatosGrid = False
        txtComentario.SetFocus
        Exit Function
    End If
    If grdPersona.TextMatrix(grdPersona.Row, 3) = "" Then
        MsgBox "Ingrese el Estado", vbInformation, "Aviso"
        ValidaDatosGrid = False
        Exit Function
    End If
    If grdPersona.TextMatrix(grdPersona.Row, 4) = "" Then
        MsgBox "Ingrese el Comentario", vbInformation, "Aviso"
        ValidaDatosGrid = False
        Exit Function
    End If


'COMENTADO BY JACA 20110609
'Dim dFecha As Date
'Dim sEstado As String, sFlag As String
'Dim i As Integer

'Valida los Nuevos datos de las persona
'For i = 1 To grdPersona.Rows - 1
'    sFlag = grdPersona.TextMatrix(i, 5)
'    If sFlag = "N" Then
'        If grdPersona.TextMatrix(i, 1) = "" Or grdPersona.TextMatrix(i, 2) = "" _
'            Or grdPersona.TextMatrix(i, 3) = "" Or grdPersona.TextMatrix(i, 4) = "" Then
'            MsgBox "Datos ingresados no válidos", vbInformation, "Aviso"
'            grdPersona.Row = i
'            ValidaDatosGrid = False
'            Exit Function
'        End If
'    End If
'Next i
'For i = 1 To grdHistoria.Rows - 1
'    sFlag = grdHistoria.TextMatrix(i, 5)
'    If sFlag = "N" Or sFlag = "M" Then
'        If grdHistoria.TextMatrix(i, 1) = "" Or grdHistoria.TextMatrix(i, 2) = "" _
'            Or grdHistoria.TextMatrix(i, 3) = "" Or grdHistoria.TextMatrix(i, 4) = "" Then
'            MsgBox "Datos ingresados no válidos", vbInformation, "Aviso"
'            grdHistoria.Row = i
'            ValidaDatosGrid = False
'            Exit Function
'        End If
'    End If
'Next i

End Function

Private Sub cmdAgregarPers_Click()

grdPersona.SetFocus
SendKeys "{Enter}"
grdPersona.TextMatrix(grdPersona.Rows - 1, 5) = "N"
cmdGrabar.Enabled = True
cmdCancelar.Enabled = True
cmdAgregarPers.Enabled = False
End Sub

Private Sub cmdBuscar_Click()
    
    Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
    Dim rsPers As ADODB.Recordset
    Dim oPers As COMDpersona.UCOMPersona
    
    dcEstado.Visible = False
    txtComentario.Visible = False
    Set oPers = frmBuscaPersona.Inicio()
    If Not oPers Is Nothing Then
        Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
        Set rsPers = clsServ.GetPersonasExoLavDinero(oPers.sPersCod)
        
        If Not (rsPers.EOF And rsPers.BOF) Then
            If nBuscar = 0 Then
                Set rsGrillaTemp = grdPersona.Recordset
            End If
            Set grdPersona.Recordset = rsPers
            Set Me.grdHistoria.Recordset = clsServ.GetPersonaHistExoLavDinero(oPers.sPersCod)
            
            nModificar = 0
            nBuscar = 1
            grdPersona.Row = 1
            lnFila = 1
            ColoreaCelda &HC0FFC0, vbBlack, grdPersona.Col
            cmdCancelar.Enabled = True
            cmdModificar.Enabled = True
            
        Else
            MsgBox "La Persona No se Encuentra en la Lista", vbInformation, "Aviso"
        End If
    End If
        Set clsServ = Nothing
End Sub

Private Sub cmdCancelar_Click()
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
    cmdModificar.Enabled = False
    dcEstado.Visible = False
    txtComentario.Visible = False
    
    cmdNuevoEst.Enabled = True
    cmdComentario.Enabled = True
    cmdAgregarPers.Enabled = True
    cmdexaminar.Enabled = True
    
    If nModificar = 2 Then
        Set grdPersona.Recordset = rsGrillaTemp
        Set rsGrillaTemp = Nothing
    ElseIf nModificar = 1 Then
        grdPersona.TextMatrix(lnFila, 3) = lsEstado
        grdPersona.TextMatrix(lnFila, 4) = lsComentario
    End If
    
    If nBuscar = 1 Then
        Set grdPersona.Recordset = rsGrillaTemp
        Set rsGrillaTemp = Nothing
        grdHistoria.Clear
        grdHistoria.Rows = 2
        grdHistoria.FormaCabecera
    End If
    nModificar = 0
    nBuscar = 0
    grdPersona.Row = lnFila
    ColoreaCelda vbWhite, vbBlack, Me.grdPersona.Col
    cmdexaminar.SetFocus
     
End Sub

Private Sub ObtienePersonas()
Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
Dim rsPers As ADODB.Recordset
Dim i As Integer
Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios

Set rsPers = clsServ.GetPersonasExoLavDinero("")
If Not (rsPers.EOF And rsPers.BOF) Then
    Set grdPersona.Recordset = rsPers
    'grdPersona_OnRowChange 1, 1
End If

Set clsServ = Nothing
End Sub

Private Sub cmdComentario_Click()
    cmdGrabar.Enabled = True
    cmdCancelar.Enabled = True
    cmdNuevoEst.Enabled = False
    cmdComentario.Enabled = False
    cmdAgregarPers.Enabled = False
    grdHistoria.Col = 3
    grdHistoria.SetFocus
    SendKeys "{Enter}"
    grdHistoria.TextMatrix(grdHistoria.Rows - 1, 6) = "M"
End Sub

Private Sub cmdexaminar_Click()
    nModificar = 2
    Dim i As Integer
    Dim oPers As COMDpersona.UCOMPersona
    Dim rsGrilla As Recordset
    
    Set oPers = frmBuscaPersona.Inicio()
    If Not oPers Is Nothing Then
        With grdPersona

            If BuscarPersonaGrilla(oPers.sPersCod) = True Then
                MsgBox "El Cliente ya se Encuentra Registrado,Verificar!", vbInformation, "Aviso"
                Exit Sub
             End If
                       
            Set rsGrilla = New Recordset
            Set rsGrillaTemp = New Recordset
             rsGrilla.Fields.Append "Codigo", adVarChar, 25
             rsGrilla.Fields.Append "Nombre", adVarChar, 300
             rsGrilla.Fields.Append "Estado", adVarChar, 150
             rsGrilla.Fields.Append "Comentario", adVarChar, 500
             rsGrilla.Open
            
             rsGrilla.AddNew
             rsGrilla.Fields("Codigo") = oPers.sPersCod
             rsGrilla.Fields("Nombre") = oPers.sPersNombre
             rsGrilla.Fields("Estado") = ""
             rsGrilla.Fields("Comentario") = ""
           
            If .Rows >= 2 And .TextMatrix(1, 1) <> "" Then
                Set rsGrillaTemp = .Recordset
                For i = 1 To .Rows - 1
                    rsGrilla.AddNew
                    rsGrilla.Fields("Codigo") = .TextMatrix(i, 1)
                    rsGrilla.Fields("Nombre") = .TextMatrix(i, 2)
                    rsGrilla.Fields("Estado") = .TextMatrix(i, 3)
                    rsGrilla.Fields("Comentario") = .TextMatrix(i, 4)
                  
                Next i
            End If
            Set .Recordset = rsGrilla
            Set rsGrilla = Nothing
            
            .Row = 1
            lnFila = 1
            ColoreaCelda &HC0FFC0, vbBlack, .Col
             .Col = 3
                dcEstado.Width = .CellWidth
                dcEstado.Left = .CellLeft + .Left
                dcEstado.Top = .CellTop + .Top
                dcEstado.Visible = True
            .Col = 4
                txtComentario.Width = .CellWidth
                txtComentario.Left = .CellLeft + .Left
                txtComentario.Top = .CellTop + .Top
                txtComentario.Visible = True
                
            grdHistoria.Clear
            grdHistoria.Rows = 2
            grdHistoria.FormaCabecera
            
            Me.cmdGrabar.Enabled = True
            Me.cmdCancelar.Enabled = True
            Me.cmdexaminar.Enabled = False
            Me.cmdModificar.Enabled = False
            dcEstado.SetFocus
        End With
    End If
End Sub
Private Function BuscarPersonaGrilla(ByVal psPersCod As String) As Boolean
    
    BuscarPersonaGrilla = False
    Dim i As Integer
    With grdPersona
        If .Rows >= 2 And .TextMatrix(1, 1) <> "" Then
            
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) = psPersCod Then
                    BuscarPersonaGrilla = True
                    Exit For
                End If
            Next i
            
        End If
    End With
End Function
Private Sub cmdGrabar_Click()
    If Not ValidaDatosGrid Then Exit Sub
    
    If MsgBox("¿Desea grabar la información?", vbQuestion + vbYesNo, "Grabar") = vbYes Then
        Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
        Dim sMovNro As String
        Dim clsMov As COMNContabilidad.NCOMContFunciones
        'Dim rsPers As ADODB.Recordset, rsHist As ADODB.Recordset ' comentado by JACA 20110609
        
        Set clsMov = New COMNContabilidad.NCOMContFunciones
        sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set clsMov = Nothing
        'Set rsPers = grdPersona.GetRsNew()' comentado by JACA 20110609
        'Set rsHist = grdHistoria.GetRsNew()' comentado by JACA 20110609
        
        Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
        With grdPersona
            clsServ.ActualizaPersExoLavDinero Trim(.TextMatrix(.Row, 1)), Right(.TextMatrix(.Row, 3), 1), Trim(.TextMatrix(.Row, 4)), sMovNro
        End With
        Set clsServ = Nothing
        ClearScreen
    End If
End Sub

Private Sub cmdModificar_Click()
    nModificar = 1
    
    lnFila = grdPersona.Row
    cmdCancelar.Enabled = True
    cmdModificar.Enabled = False
    Me.cmdexaminar.Enabled = False
    cmdGrabar.Enabled = True
    
    With grdPersona
         .Col = 3
            lsEstado = dcEstado.Text + Space(75) + dcEstado.BoundText
            dcEstado.Width = .CellWidth
            dcEstado.Left = .CellLeft + .Left
            dcEstado.Top = .CellTop + .Top
            dcEstado.Visible = True
        .Col = 4
            lsComentario = .TextMatrix(.Row, 4)
            txtComentario.Text = .TextMatrix(.Row, 4)
            txtComentario.Width = .CellWidth
            txtComentario.Left = .CellLeft + .Left
            txtComentario.Top = .CellTop + .Top
            txtComentario.Visible = True
        .Col = 3
        dcEstado.SetFocus
        
    End With
End Sub

Private Sub cmdNuevoEst_Click()
    cmdGrabar.Enabled = True
    cmdCancelar.Enabled = True
    cmdNuevoEst.Enabled = False
    cmdComentario.Enabled = False
    cmdAgregarPers.Enabled = False
    grdHistoria.AdicionaFila
    grdHistoria.SetFocus
    SendKeys "{Enter}"
    grdHistoria.TextMatrix(grdHistoria.Rows - 1, 1) = Format$(gdFecSis, gcFormatoFechaView)
    grdHistoria.TextMatrix(grdHistoria.Rows - 1, 6) = "N"
    grdHistoria.TextMatrix(grdHistoria.Rows - 1, 4) = grdPersona.TextMatrix(grdPersona.Row, 1)
End Sub



Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
Me.Caption = "Registro Personas Exoneradas Lavado Dinero"
nModificar = 0
CargarCabeceraPersona
CargarComboEstado
CargarCombos
ObtienePersonas
'Dim rsEst As ADODB.Recordset
'Dim rsHist As ADODB.Recordset
'Set rsHist = New ADODB.Recordset

'Dim clsGen As COMDConstSistema.DCOMGeneral
'Set clsGen = New COMDConstSistema.DCOMGeneral
'Set rsHist = rsEst.Clone
'grdPersona.CargaCombo rsEst
''Set rsEst = clsGen.GetConstante(gPersEstLavDinero)
'grdHistoria.CargaCombo rsHist 'rsEst
'Set clsGen = Nothing

cmdNuevoEst.Enabled = False
cmdComentario.Enabled = False
End Sub

'Private Sub grdPersona_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
'If psDataCod = "" Then
'    grdPersona.EliminaFila pnRow
'End If
'If pbEsDuplicado Then
'    MsgBox "Persona registrada. Ingrese un nuevo estado.", vbInformation, "Aviso"
'    grdPersona.EliminaFila pnRow
'    If cmdNuevoEst.Enabled Then cmdNuevoEst.SetFocus
'End If
'End Sub

'Private Sub grdPersona_OnRowChange(pnRow As Long, pnCol As Long)
'Dim rsHist As ADODB.Recordset
'Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
'Dim sPersCod As String
'
'sPersCod = grdPersona.TextMatrix(pnRow, 1)
'
'Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
'Set rsHist = clsServ.GetPersonaHistExoLavDinero(sPersCod)
'Set clsServ = Nothing
'
'If Not (rsHist.EOF And rsHist.BOF) Then
'    Set grdHistoria.Recordset = rsHist
'    cmdNuevoEst.Enabled = True
'    cmdComentario.Enabled = True
'Else
'    grdHistoria.Clear
'    grdHistoria.Rows = 2
'    grdHistoria.FormaCabecera
'    cmdNuevoEst.Enabled = False
'    cmdComentario.Enabled = False
'End If
'Set rsHist = Nothing
'End Sub
'JACA 20110608***************************************
Private Sub CargarCabeceraPersona()
    grdPersona.Clear
    grdPersona.Rows = 2
    
    With grdPersona
        .TextMatrix(0, 0) = "#"
        .TextMatrix(0, 1) = "Codigo"
        .TextMatrix(0, 2) = "Nombre"
        .TextMatrix(0, 3) = "Estado"
        '.TextMatrix(0, 4) = "Nombre" ' se modifico con NºOpe
        .TextMatrix(0, 4) = "Comentario"
        
       
        
        .ColWidth(0) = 350
        .ColWidth(1) = 1300 'Codigo
        .ColWidth(2) = 3000 'Nombre
        .ColWidth(3) = 1000  'Estado
        .ColWidth(4) = 3500  'Comentario
        .ColWidth(5) = 0  'Flag
       
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
         
        .ColAlignmentFixed(0) = flexAlignCenterCenter
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
        
    End With
End Sub

Private Sub CargarCombos()
    Dim rsEst As ADODB.Recordset
    Set rsEst = New ADODB.Recordset
    
         rsEst.Fields.Append "cDescripcion", adVarChar, 200
         rsEst.Fields.Append "nValor", adInteger
         rsEst.Open
        
         rsEst.AddNew
         rsEst.Fields("cDescripcion") = "ACTIVO"
         rsEst.Fields("nValor") = 1
         
         rsEst.AddNew
         rsEst.Fields("cDescripcion") = "INACTIVO"
         rsEst.Fields("nValor") = 0
         'rsEst.MoveFirst
         
        dcEstado.BoundColumn = "nValor"
        dcEstado.DataField = "nValor"
        Set dcEstado.RowSource = rsEst
        dcEstado.ListField = "cDescripcion"
        dcEstado.BoundText = 0
        
        
        
End Sub

Sub CargarComboEstado()
    grdPersona.RowHeightMin = dcEstado.Height
    dcEstado.Visible = False
    dcEstado.Width = grdPersona.CellWidth
End Sub
Private Sub dcEstado_Click(Area As Integer)
'With grdPersona
'    .RowSel = .Row
'    .Col = 0
'    .ColSel = .Cols - 1
'
'    .BackColorSel = &HC0FFC0
'    .ForeColorSel = vbBlack
'  End With
'    If dcCordinador.Visible = False Then
'        dcCordinador.Visible = True
'        dcCordinador.SetFocus
'    Else
'        grdAnalista.TextMatrix(grdAnalista.Row, 10) = dcCordinador.Text
'        dcCordinador.Visible = False
'        grdAnalista.Col = grdAnalista.Col + 1
'        grdAnalista.CellBackColor = &H80000018
'        grdAnalista.SetFocus
'    End If
End Sub

Private Sub dcEstado_KeyPress(KeyAscii As Integer)
    grdPersona.Col = 3
    If KeyAscii = 13 Then
            grdPersona.TextMatrix(lnFila, 3) = dcEstado.Text + Space(75) + dcEstado.BoundText
            dcEstado.Visible = False
            grdPersona.Col = grdPersona.Col + 1
            txtComentario.Text = grdPersona.TextMatrix(lnFila, 4)
            txtComentario.Width = grdPersona.CellWidth
            txtComentario.Left = grdPersona.CellLeft + grdPersona.Left
            txtComentario.Top = grdPersona.CellTop + grdPersona.Top
            txtComentario.Visible = True
            txtComentario.SetFocus
        End If
End Sub
'Private Sub cmdModificarPersona_Click()
'    Me.cmdCancelar.Enabled = True
'    Me.cmdGrabar.Enabled = True
'    Me.grdPersona.Row = Me.grdPersona.Row
'    Me.grdPersona.BackColorRow vbYellow
'End Sub

Private Sub grdPersona_LeaveCell()
    If nModificar = 0 Then
        ColoreaCelda vbWhite, vbBlack, Me.grdPersona.Col
    End If
End Sub

Private Sub grdPersona_RowColChange()
    Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
    Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
    If nModificar = 0 Then
        ColoreaCelda &HC0FFC0, vbBlack, Me.grdPersona.Col
        Me.cmdModificar.Enabled = True
        Set Me.grdHistoria.Recordset = clsServ.GetPersonaHistExoLavDinero(grdPersona.TextMatrix(grdPersona.Row, 1))
        
    End If
End Sub
Private Sub grdPersona_Click()
  Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
  Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
  With grdPersona
    
     If nModificar = 0 Then
        Me.cmdModificar.Enabled = True
        ColoreaCelda &HC0FFC0, vbBlack, .Col
        Set Me.grdHistoria.Recordset = clsServ.GetPersonaHistExoLavDinero(.TextMatrix(.Row, 1))
        'GetPersonaHistExoLavDinero
     End If
    
    If .Col = 3 And .Row = lnFila And nModificar > 0 Then
        dcEstado.Width = .CellWidth
        dcEstado.Left = .CellLeft + .Left
        dcEstado.Top = .CellTop + .Top
        dcEstado.Visible = True
    ElseIf .Col = 4 And .Row = lnFila And nModificar > 0 Then
        dcEstado.Visible = False
        txtComentario.Text = .TextMatrix(.Row, 4)
        txtComentario.Width = .CellWidth
        txtComentario.Left = .CellLeft + .Left
        txtComentario.Top = .CellTop + .Top
        txtComentario.Visible = True
    Else
        dcEstado.Visible = False
        txtComentario.Visible = False
    End If
     
  End With
End Sub
Private Sub ColoreaCelda(ByVal colorCelda As OLE_COLOR, ByVal colorFuente As OLE_COLOR, ByVal nCol As Integer)
    lnColumna = nCol
    With grdPersona
        Dim i As Integer
        For i = 1 To .Cols - 1
            .Col = i
            .CellBackColor = colorCelda
            .CellForeColor = colorFuente
        Next i
        .Col = lnColumna
    End With
    
End Sub
Private Sub txtComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grdPersona.TextMatrix(lnFila, 4) = Trim(txtComentario.Text)
        txtComentario.Visible = False
        Me.cmdGrabar.Enabled = True
        Me.txtComentario.Text = ""
        Me.cmdGrabar.SetFocus
    End If
End Sub
'JACA END********************************************


