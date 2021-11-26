VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmInvActivarBienDolares 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Activar Bien como Activo Fijo Dolares"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13485
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvActivarBienDolares.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   13485
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraRefrescar 
      Appearance      =   0  'Flat
      Caption         =   "Parámetros de Búsqueda"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   135
      TabIndex        =   3
      Top             =   120
      Width           =   6120
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4560
         TabIndex        =   4
         Top             =   360
         Width           =   1230
      End
      Begin MSMask.MaskEdBox mskIni 
         Height          =   300
         Left            =   600
         TabIndex        =   5
         Top             =   360
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
      Begin MSMask.MaskEdBox mskFin 
         Height          =   300
         Left            =   3000
         TabIndex        =   6
         Top             =   360
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
      Begin VB.Label lblFinal 
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   2280
         TabIndex        =   8
         Top             =   360
         Width           =   510
      End
      Begin VB.Label lblInicial 
         Caption         =   "De:"
         Height          =   195
         Left            =   210
         TabIndex        =   7
         Top             =   360
         Width           =   285
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Visualizar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   2
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "LISTA DE ACTIVOS FIJOS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   13095
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fg 
         Height          =   4215
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   7435
         _Version        =   393216
         Rows            =   3
         Cols            =   12
         FixedRows       =   2
         BackColorBkg    =   -2147483643
         AllowBigSelection=   0   'False
         TextStyleFixed  =   3
         FocusRect       =   2
         GridLinesFixed  =   1
         GridLinesUnpopulated=   3
         Appearance      =   0
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
         _NumberOfBands  =   1
         _Band(0).Cols   =   12
      End
   End
End
Attribute VB_Name = "frmInvActivarBienDolares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmInvActivarBienDolares
'** Descripción : Formulario para Listar los BS filtrados por Activo Fijo en Dolares
'** Creación : MAVM, 20090218 8:59:25 AM
'** Modificación:
'********************************************************************

Option Explicit

Private Sub Form_Load()
    CentraForm Me
    Me.mskIni.Text = "01/01/" & Format(gdFecSis, "yyyy")
    Me.mskFin.Text = Format(gdFecSis, gsFormatoFechaView)
    CargarCabecera
End Sub

Private Sub Command1_Click()
    gsMvoNro = fg.TextMatrix(fg.Row, 8)
    gsMonedaAF = "2"
    frmInvRegistrarAF.LLenarDatos
End Sub

Private Sub CargarCabecera()
    fg.Cols = 9
    fg.TextMatrix(0, 0) = " "
    fg.TextMatrix(1, 0) = " "
    fg.TextMatrix(0, 1) = "Documento"
    fg.TextMatrix(0, 2) = "Documento"
    fg.TextMatrix(0, 3) = "Documento"
    fg.TextMatrix(0, 4) = "Documento"
    
    fg.TextMatrix(1, 1) = "Tipo"
    fg.TextMatrix(1, 2) = "Número"
    
    fg.TextMatrix(1, 3) = "NºOC Di/Pr"
    fg.TextMatrix(1, 4) = "Fecha"
    
    fg.TextMatrix(0, 5) = "Proveedor"
    fg.TextMatrix(1, 5) = "Proveedor"
    
    fg.TextMatrix(0, 6) = "Importe"
    fg.TextMatrix(1, 6) = "Importe"
    fg.TextMatrix(0, 7) = "Observaciones"
    fg.TextMatrix(1, 7) = "Observaciones"
    fg.TextMatrix(1, 8) = "cMovNro"
'    fg.TextMatrix(1, 9) = "nImporte"
'
'    fg.TextMatrix(1, 10) = "Saldo"
'    fg.TextMatrix(1, 11) = "Estado"
'    fg.TextMatrix(1, 12) = "Monto ($)"
    
    fg.RowHeight(-1) = 285
    fg.ColWidth(0) = 400
    fg.ColWidth(1) = 500
    fg.ColWidth(2) = 1200
    fg.ColWidth(3) = 1200
    fg.ColWidth(4) = 1000
    
    fg.ColWidth(5) = 3200
    fg.ColWidth(6) = 1200
    fg.ColWidth(7) = 3870
    fg.ColWidth(8) = 0
    fg.ColWidth(9) = 0
'    If lbPresu Then
'       fg.ColWidth(10) = 0
'       fg.ColWidth(11) = 0
'    Else
'       fg.ColWidth(10) = 0 '1200
'       fg.ColWidth(11) = 0 '1700
'    End If
    
'    If lbReporteFechas Then
'        fg.ColWidth(12) = 0 '1700
'    Else
'        fg.ColWidth(12) = 0
'    End If
    
    fg.MergeCells = flexMergeRestrictColumns
    fg.MergeCol(0) = True
    fg.MergeCol(1) = True
    fg.MergeCol(2) = True
    fg.MergeCol(3) = True
    fg.MergeCol(4) = True
    fg.MergeCol(5) = True
    fg.MergeCol(6) = True
    fg.MergeCol(7) = True
    
    fg.MergeRow(0) = True
    fg.MergeRow(1) = True
    fg.RowHeight(0) = 200
    fg.RowHeight(1) = 200
    fg.ColAlignmentFixed(-1) = flexAlignCenterCenter
    fg.ColAlignment(1) = flexAlignCenterCenter
    fg.ColAlignment(3) = flexAlignCenterCenter
    fg.ColAlignment(6) = flexAlignLeftCenter
End Sub

Private Sub mskIni_GotFocus()
    mskIni.SelStart = 0
    mskIni.SelLength = 50
End Sub

Private Sub mskFin_GotFocus()
    mskFin.SelStart = 0
    mskFin.SelLength = 50
End Sub

Private Sub mskIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.mskFin.SetFocus
    End If
End Sub

Private Sub mskFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdActualizar.SetFocus
    End If
End Sub

Private Sub cmdActualizar_Click()
    CargarDatos
End Sub

Private Sub CargarDatos()
    Dim rsCargar As ADODB.Recordset
    Dim oInventario As NInvActivoFijo
    Set oInventario = New NInvActivoFijo
    Dim nItem, nTot As Integer
    
    Set rsCargar = oInventario.DarOrdenCompraXAF(Format(mskIni.Text, "yyyymmdd"), Format(mskFin.Text, "yyyymmdd"), "252601")
    
    fg.Rows = 3
    nItem = 1
    nTot = 0
    Do While Not rsCargar.EOF
       If nItem <> 1 Then
          AdicionaRow fg
       End If
       nItem = fg.Row
       fg.TextMatrix(nItem, 0) = nItem - 1
       fg.TextMatrix(nItem, 1) = rsCargar!cDocAbrev
       fg.TextMatrix(nItem, 2) = rsCargar!cDocNro
    
       fg.TextMatrix(nItem, 3) = IIf(IsNull(rsCargar!tipoOc), "", rsCargar!tipoOc) + "-" + rsCargar!cdocnroOCD
    
       fg.TextMatrix(nItem, 4) = rsCargar!dDocFecha
       If Not IsNull(rsCargar!cNomPers) Then
          fg.TextMatrix(nItem, 5) = PstaNombre(Trim(Mid(rsCargar!cNomPers, 1, Len(rsCargar!cNomPers) - 50)), True)
       End If
       fg.TextMatrix(nItem, 6) = Format(rsCargar!nDocImporte, gcFormView)
       fg.TextMatrix(nItem, 7) = rsCargar!cMovDesc
       fg.TextMatrix(nItem, 8) = rsCargar!nMovNro
'       fg.TextMatrix(nItem, 9) = Right(rsCargar!cNomPers, 13) & "" ' rsCargar!cBSCod 'CODIGO PERSONA
'       fg.TextMatrix(nItem, 10) = Format(rsCargar!nDocImporte, gcFormView)
'       If rsCargar!nMovEstado = 16 And Not rsCargar!nMovFlag = gMovFlagExtornado Then
'          fg.TextMatrix(nItem, 11) = "Aprobado"
'       ElseIf rsCargar!nMovEstado = 15 And Not rsCargar!nMovFlag = gMovFlagExtornado Then
'          fg.TextMatrix(nItem, 11) = "Pendiente"
'          fg.Col = 11
'          fg.CellBackColor = "&H00C0C0FF"
'       ElseIf rsCargar!nMovEstado = 14 Then
'          fg.TextMatrix(nItem, 11) = "RECHAZADO"
'          fg.Col = 11
'          fg.CellBackColor = "&H0080FF80"
'       Else
'          If rsCargar!nMovFlag = gMovFlagExtornado Or rsCargar!nMovFlag = gMovFlagEliminado Then
'             fg.TextMatrix(nItem, 11) = "ELIMINADO"
'             fg.Col = 11
'             fg.CellBackColor = "&H0080FF80"
'          End If
'       End If
'       'If lbReporteFechas Then fg.TextMatrix(nItem, 12) = Format(rsCargar!nDocMEImporte, gcFormView)
'
'       nTot = nTot + rsCargar!nDocImporte
       rsCargar.MoveNext
    Loop
    'txtTot = Format(nTot, gcFormView)
    fg.Row = 2
    fg.Col = 1
    Set rsCargar = Nothing
    Set oInventario = Nothing
End Sub

