VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCapBeneficiario 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   Icon            =   "frmCapBeneficiario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraBenef 
      Caption         =   "Beneficiarios"
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
      Height          =   3135
      Left            =   105
      TabIndex        =   10
      Top             =   1875
      Width           =   8475
      Begin SICMACT.FlexEdit grdBenef 
         Height          =   2325
         Left            =   105
         TabIndex        =   2
         Top             =   210
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   4101
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Codigo-Nombre-Parentesco-Edad-%-CodParent"
         EncabezadosAnchos=   "350-1400-3500-1500-600-700-0"
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
         ColumnasAEditar =   "X-1-X-3-X-5-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-1-0-3-0-0-0"
         EncabezadosAlineacion=   "C-L-L-L-C-R-C"
         FormatosEdit    =   "0-0-0-0-0-2-2"
         TextArray0      =   "#"
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   6255
         TabIndex        =   3
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   7335
         TabIndex        =   4
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2625
         TabIndex        =   13
         Top             =   2730
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "Porcentaje Acumulado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   12
         Top             =   2625
         Width           =   1245
      End
      Begin VB.Label lblPorcentaje 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   1455
         TabIndex        =   11
         Top             =   2655
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   5115
      Width           =   915
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1050
      TabIndex        =   6
      Top             =   5115
      Width           =   915
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7665
      TabIndex        =   8
      Top             =   5115
      Width           =   915
   End
   Begin VB.Frame fraAsegurado 
      Caption         =   "Asegurado"
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
      Height          =   1770
      Left            =   105
      TabIndex        =   9
      Top             =   0
      Width           =   8475
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCliente 
         Height          =   1380
         Left            =   105
         TabIndex        =   1
         Top             =   210
         Width           =   7260
         _ExtentX        =   12806
         _ExtentY        =   2434
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   7455
         TabIndex        =   0
         Top             =   210
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   6615
      TabIndex        =   7
      Top             =   5115
      Width           =   915
   End
End
Attribute VB_Name = "frmCapBeneficiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bConsulta As Boolean
Dim nPorcAcum As Double
'By capi 21012009
Dim objPista As COMManejador.Pista



Public Sub SetupGridCliente()
Dim i As Integer
For i = 1 To grdCliente.Rows - 1
    grdCliente.MergeCol(i) = True
Next i
grdCliente.MergeCells = flexMergeFree
grdCliente.BandExpandable(0) = True
grdCliente.Cols = 9
grdCliente.ColWidth(0) = 100
grdCliente.ColWidth(1) = 3500
grdCliente.ColWidth(2) = 3500
grdCliente.ColWidth(3) = 1500
grdCliente.ColWidth(4) = 1000
grdCliente.ColWidth(5) = 600
grdCliente.ColWidth(6) = 1500
grdCliente.TextMatrix(0, 1) = "Nombre"
grdCliente.TextMatrix(0, 2) = "Dirección"
grdCliente.TextMatrix(0, 3) = "Zona"
grdCliente.TextMatrix(0, 4) = "Fono"
grdCliente.TextMatrix(0, 5) = "ID"
grdCliente.TextMatrix(0, 6) = "ID N°"
End Sub
        
Private Sub PorcentajeAcumulado()
Dim i As Integer, nFila As Long, nCol As Long
Dim nAcum As Double
nAcum = 0
nFila = grdBenef.Row
nCol = grdBenef.Col
For i = 1 To grdBenef.Rows - 1
    If grdBenef.TextMatrix(i, 5) <> "" Then
        nAcum = nAcum + CDbl(grdBenef.TextMatrix(i, 5))
    End If
Next i
grdBenef.Row = nFila
grdBenef.Col = nCol
nPorcAcum = nAcum
If nPorcAcum = 100 Then
    cmdAgregar.Enabled = False
ElseIf nPorcAcum < 100 Then
    cmdAgregar.Enabled = True
ElseIf nPorcAcum = 0 Then
    cmdEliminar.Enabled = False
    cmdAgregar.SetFocus
End If
lblPorcentaje = Format$(nPorcAcum, "#,##0.00")
End Sub

Public Sub Inicia(Optional bCons As Boolean = True)
bConsulta = bCons
If bConsulta Then
    cmdGrabar.Visible = False
    cmdCancelar.Visible = False
    cmdEliminar.Visible = False
    cmdAgregar.Visible = False
    grdBenef.lbEditarFlex = False
    grdBenef.VisiblePopMenu = False
Else
    cmdGrabar.Visible = True
    cmdCancelar.Visible = True
    cmdEliminar.Visible = True
    cmdAgregar.Visible = True
    grdBenef.lbEditarFlex = True
    grdBenef.VisiblePopMenu = True
    
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim rsPar As ADODB.Recordset
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsPar = clsGen.GetConstante(gPersRelacion)
    grdBenef.CargaCombo rsPar
    Set clsGen = Nothing
End If
SetupGridCliente
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
fraBenef.Enabled = False
lblPorcentaje = "0.00"
Me.Show 1
End Sub

Private Sub CmdAgregar_Click()
If nPorcAcum < 100 Then
    grdBenef.AdicionaFila
    grdBenef.TextMatrix(grdBenef.Rows - 1, 5) = "0.00"
    grdBenef.SetFocus
    SendKeys "{ENTER}"
Else
    MsgBox "No es posible agregar más beneficiarios si el porcentaje acumulado es 100%", vbInformation, "Aviso"
End If
End Sub

Private Sub cmdBuscar_Click()
Dim clsPers As COMDPersona.UCOMPersona
Dim nPersoneria As COMDConstantes.PersPersoneria

Set clsPers = New COMDPersona.UCOMPersona
Set clsPers = frmBuscaPersona.Inicio

If Not clsPers Is Nothing Then
    Dim sPers As String
    Dim rsPers As New ADODB.Recordset
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaGenerales
    
    
    nPersoneria = CLng(clsPers.sPersPersoneria)
    
    If nPersoneria = gPersonaNat Then
        sPers = clsPers.sPerscod
        Set clsPers = Nothing
        Set clsCap = New COMNCaptaGenerales.NCOMCaptaGenerales
        Set rsPers = clsCap.GetDatosPersona(sPers)
        If Not (rsPers.EOF And rsPers.EOF) Then
            Set grdCliente.Recordset = rsPers
            SetupGridCliente
        End If
        
        grdBenef.Clear
        grdBenef.FormaCabecera
        grdBenef.Rows = 2
        Set rsPers = clsCap.GetBeneficiarios(sPers)
        Set clsCap = Nothing
        If Not (rsPers.EOF And rsPers.EOF) Then
            fraBenef.Enabled = True
            Set grdBenef.Recordset = rsPers
            If Not bConsulta Then
                cmdAgregar.Enabled = True
                cmdAgregar.SetFocus
                cmdGrabar.Enabled = True
                cmdCancelar.Enabled = True
            End If
            cmdPrint.Enabled = True
            PorcentajeAcumulado
        Else
            If bConsulta Then
                MsgBox "Persona no posee ningun beneficiario registrado.", vbInformation, "Aviso"
                fraBenef.Enabled = False
                cmdPrint.Enabled = False
                cmdBuscar.SetFocus
            Else
                fraBenef.Enabled = True
                cmdAgregar.Enabled = True
                cmdAgregar.SetFocus
                cmdGrabar.Enabled = True
                cmdCancelar.Enabled = True
                cmdPrint.Enabled = False
            End If
        End If
        rsPers.Close
        Set rsPers = Nothing
    Else
        MsgBox "El Asegurado sólo puede ser una PERSONA NATURAL", vbInformation, "Aviso"
        cmdBuscar.SetFocus
    End If
Else
    cmdBuscar.SetFocus
End If

End Sub

Private Sub cmdCancelar_Click()
grdCliente.Clear
grdCliente.Rows = 2
SetupGridCliente
grdBenef.Clear
grdBenef.Rows = 2
grdBenef.FormaCabecera
fraBenef.Enabled = False
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
cmdBuscar.SetFocus
nPorcAcum = 0
lblPorcentaje = "0.00"
End Sub

Private Sub cmdEliminar_Click()
If grdBenef.TextMatrix(1, 1) <> "" Then
    Dim nItem As Long
    nItem = grdBenef.Row
    If MsgBox("Está seguro de eliminar al beneficiario??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        grdBenef.EliminaFila nItem
        PorcentajeAcumulado
    End If
Else
    MsgBox "No existen beneficiarios que eliminar", vbInformation, "Aviso"
End If
End Sub

Private Sub CmdGrabar_Click()
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim sPersona As String
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
sPersona = grdCliente.TextMatrix(1, 7)
If nPorcAcum = 0 Then
    If MsgBox("No ha registrado ningún beneficiario. ¿Desea eliminar a sus beneficiarios anteriores?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        clsMant.EliminaBeneficiarios sPersona
    Else
        cmdAgregar.SetFocus
    End If
    Exit Sub
End If
If MsgBox("¿Está seguro que desea grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim rsBenef As New ADODB.Recordset
    Dim sMovNro As String
    Dim clsMov As COMNContabilidad.NCOMContFunciones
    Set clsMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = clsMov.GeneraMovNro(gdFecSis, , gsCodUser)
    Set clsMov = Nothing
    Set rsBenef = grdBenef.GetRsNew
    If clsMant.ActualizaBeneficiarios(sPersona, rsBenef, sMovNro) Then
        cmdCancelar_Click
    End If
    'By Capi 21012009
    objPista.InsertarPista gsOpeCod, sMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, , sPersona, gCodigoPersona
    '
    
End If
Set clsMant = Nothing
End Sub

Private Sub cmdPrint_Click()
    Dim sNomCli As String, sCodCli As String
    Dim sDirCli As String, sDocId As String
    Dim clsPrev As previo.clsprevio
    Dim rs As New ADODB.Recordset
    Dim lsCaImp As String
    Dim oCapI As COMNCaptaGenerales.NCOMCaptaImpresion
    Dim i As Integer
    
    sNomCli = grdCliente.TextMatrix(1, 1)
    sDirCli = grdCliente.TextMatrix(1, 2)
    sCodCli = grdCliente.TextMatrix(1, 5)
    sDocId = grdCliente.TextMatrix(1, 4)
        
   
    
    With rs
            'Crear RecordSet
        
        
        .Fields.Append "cVal1", adVarChar, 150
        .Fields.Append "cNomCli", adVarChar, 150
        .Fields.Append "cParCli", adVarChar, 150
        .Fields.Append "cVal2", adVarChar, 150
        .Fields.Append "cVal3", adVarChar, 150
        .Open
        'Llenar Recordset
        For i = 1 To grdBenef.Rows - 1
             .AddNew
            .Fields("cVal1") = grdBenef.TextMatrix(i, 1)
            .Fields("cNomCli") = PstaNombre(grdBenef.TextMatrix(i, 2))
            .Fields("cParCli") = UCase(Trim(Left(grdBenef.TextMatrix(i, 3), 25)))
            .Fields("cVal2") = grdBenef.TextMatrix(i, 4)
            .Fields("cVal3") = grdBenef.TextMatrix(i, 5)
        Next i
        
    End With
    Set oCapI = New COMNCaptaGenerales.NCOMCaptaImpresion
        lsCaImp = oCapI.ImprimeBeneficiario(gsNomAge, gdFecSis, sNomCli, sDirCli, sCodCli, sDocId, nPorcAcum, rs)
    Set oCapI = Nothing
    Set clsPrev = New previo.clsprevio
        clsPrev.Show lsCaImp, "Declaración de Beneficiarios", False, 66, gImpresora
    Set clsPrev = Nothing

End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
Me.Caption = "Declaración de Beneficiarios"
'By Capi 20012009
Set objPista = New COMManejador.Pista
gsOpeCod = gCapMantBeneficiarios
'End By
End Sub

Private Sub grdBenef_OnCellChange(pnRow As Long, pnCol As Long)
If pnCol = 5 Then
    If grdBenef.TextMatrix(pnRow, pnCol) <> "" Then
        Dim nMonto As Double
        nMonto = CDbl(grdBenef.TextMatrix(pnRow, pnCol))
        If nMonto < 0 Then
            grdBenef.TextMatrix(pnRow, pnCol) = "0.00"
        Else
            If nPorcAcum + nMonto <= 100 Then
                PorcentajeAcumulado
            Else
                MsgBox "Porcentaje sobrepasa al permitido. Solo puede otorgar como máximo un " & (100 - nPorcAcum) & " %", vbInformation, "Aviso"
                grdBenef.TextMatrix(pnRow, pnCol) = Format$(100 - nPorcAcum, "#,##0.00")
            End If
        End If
    End If
End If
End Sub

Private Sub grdBenef_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
Dim sPersona As String
sPersona = grdCliente.TextMatrix(1, 7)
If sPersona = psDataCod Then
    MsgBox "El Beneficiario no puede ser EL MISMSO ASEGURADO", vbInformation, "Aviso"
    grdBenef.EliminaFila pnRow
    cmdAgregar.SetFocus
    Exit Sub
End If
If Not pbEsDuplicado Then
    Dim nPersoneria As COMDConstantes.PersPersoneria
    nPersoneria = grdBenef.PersPersoneria
    If nPersoneria = gPersonaNat Then
        Dim clsGen As COMDConstSistema.DCOMGeneral
        Dim nEdad As Long
        Set clsGen = New COMDConstSistema.DCOMGeneral
        nEdad = clsGen.GetPersonaEdad(psDataCod)
        Set clsGen = Nothing
        grdBenef.TextMatrix(pnRow, 4) = nEdad
    Else
        MsgBox "El Beneficiario sólo puede ser PERSONA NATURAL", vbInformation, "Aviso"
        grdBenef.EliminaFila pnRow
        cmdAgregar.SetFocus
    End If
End If


End Sub

