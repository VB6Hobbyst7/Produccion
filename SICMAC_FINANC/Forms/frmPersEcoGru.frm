VERSION 5.00
Begin VB.Form frmPersEcoGru 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Grupos Economicos"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8580
   Icon            =   "frmPersEcoGru.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   345
      Left            =   60
      TabIndex        =   3
      Top             =   3885
      Width           =   1005
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   345
      Left            =   1110
      TabIndex        =   2
      Top             =   3885
      Width           =   1005
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   7530
      TabIndex        =   1
      Top             =   3885
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Top             =   3885
      Width           =   1005
   End
   Begin VB.Frame fraGE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   3810
      Left            =   45
      TabIndex        =   4
      Top             =   30
      Width           =   8520
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   345
         Left            =   6390
         TabIndex        =   7
         Top             =   3390
         Width           =   1005
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   345
         Left            =   7440
         TabIndex        =   6
         Top             =   3390
         Width           =   1005
      End
      Begin Sicmact.FlexEdit Flex 
         Height          =   3075
         Left            =   105
         TabIndex        =   5
         Top             =   240
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   5424
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Codigo-Descripcion-Tipo"
         EncabezadosAnchos=   "400-1500-4000-2000"
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
         ColumnasAEditar =   "X-X-2-3"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-3"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmPersEcoGru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelar_Click()
    Form_Load
End Sub

Private Sub cmdEditar_Click()
    Activa True
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Desea Eliminar el registro seleccionado ? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Flex.EliminaFila Flex.Row
End Sub

Private Sub cmdGrabar_Click()
    Dim oGE As DGrupoEco
    Set oGE = New DGrupoEco
    
    If Not Valida Then Exit Sub
    
    oGE.ActulizaGE Flex.GetRsNew
    
    Activa False
End Sub

Private Sub cmdNuevo_Click()
    Flex.AdicionaFila
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Flex_OnRowAdd(pnRow As Long)
    If pnRow = 1 And Flex.TextMatrix(pnRow, 1) = "" Then
        Flex.TextMatrix(pnRow, 1) = "000001"
    Else
        Flex.TextMatrix(pnRow, 1) = Format(CCur(Flex.TextMatrix(pnRow - 1, 1)) + 1, "000000")
    End If
End Sub

Private Sub Form_Load()
    Dim oCon As DConstantes
    Set oCon = New DConstantes
    Dim oGE As DGrupoEco
    
    Me.Icon = LoadPicture(App.path & "\Graficos\cm.ico")
    Set oGE = New DGrupoEco
    Flex.CargaCombo oCon.GetConstante(4027, , , , , , True)
    
    Flex.rsFlex = oGE.GetGE
    
    Set oCon = Nothing
    Set oGE = Nothing
    
    Activa False
    CentraForm Me
End Sub

Private Sub Activa(pbEdita As Boolean)
    Me.cmdEditar.Visible = Not pbEdita
    Me.cmdGrabar.Enabled = pbEdita
    Me.cmdCancelar.Visible = pbEdita
    Me.fraGE.Enabled = pbEdita
End Sub

Private Function Valida() As Boolean
    Dim lnI As Integer
    
    For lnI = 1 To Me.Flex.Rows - 1
        Flex.Row = lnI
        If Me.Flex.TextMatrix(lnI, 1) = "" Then
            MsgBox "Se ha generado erroneamente el codigo del resgistro " & lnI, vbInformation, "Aviso"
            Flex.Col = 1
            Valida = False
            Exit Function
        ElseIf Me.Flex.TextMatrix(lnI, 2) = "" Then
            MsgBox "Debe ingresar una descripcon del grupo economico." & lnI, vbInformation, "Aviso"
            Flex.Col = 2
            Valida = False
            Exit Function
        ElseIf Me.Flex.TextMatrix(lnI, 3) = "" Then
            MsgBox "Debe ingresar un tipo de grupo economico." & lnI, vbInformation, "Aviso"
            Flex.Col = 3
            Valida = False
            Exit Function
        Else
            Valida = True
        End If
    Next lnI
End Function
