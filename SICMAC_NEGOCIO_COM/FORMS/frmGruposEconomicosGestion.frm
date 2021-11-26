VERSION 5.00
Begin VB.Form frmGruposEconomicosGestion 
   Caption         =   "Registrar Vinculados - Gestión"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7935
   Icon            =   "frmGruposEconomicosGestion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   5400
      TabIndex        =   12
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6720
      TabIndex        =   11
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Grupo Economico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.Frame FraGestion 
         Caption         =   "Gestión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   0
         TabIndex        =   1
         Top             =   1440
         Width           =   7695
         Begin SICMACT.FlexEdit FEGestion 
            Height          =   3255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   5741
            Cols0           =   3
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "-Codigo-Descripcion"
            EncabezadosAnchos=   "400-1200-5700"
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
            ColumnasAEditar =   "X-1-X"
            ListaControles  =   "0-3-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L"
            FormatosEdit    =   "0-0-0"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Grupo                :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Empresa             :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Vinculado           :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblGrupoEconomico 
         Caption         =   "[Mostrar]"
         Height          =   255
         Left            =   2040
         TabIndex        =   4
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label lblEmpresaVinculado 
         Caption         =   "[Mostrar]"
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label lblPersonaVinculado 
         Caption         =   "[Mostrar]"
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Top             =   1080
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmGruposEconomicosGestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FEGestionNoMoverdeFila As Integer
Dim lnGrupo As Integer
Dim lsPersCodEmpresa As String
Dim lsPersCodVinculado As String

Private Sub cmdEliminar_Click()
If FEGestion.Rows > 0 Then
    If FEGestion.row >= 1 Then
        Call FEGestion.EliminaFila(FEGestion.row)
        'FEGestion.Rows = FEGestion.Rows - 1
    End If
End If
End Sub

Private Sub cmdGrabar_Click()
Dim i As Integer
Dim oGrup As COMDpersona.DCOMGrupoE
Dim nRetorno As Integer
Set oGrup = New COMDpersona.DCOMGrupoE
With FEGestion
    For i = 1 To .Rows - 1
    nRetorno = oGrup.ActualizarRelacionGestionVinculado(Trim(.TextMatrix(i, 1)), lsPersCodEmpresa, lsPersCodVinculado, lnGrupo, 1, i)
    Next
 MsgBox "Datos se guardaron correctamente", vbApplicationModal
Call cmdSalir_Click
End With
End Sub

Private Sub cmdNuevo_Click()
FEGestion.AdicionaFila
'JAME 20140303 ERS167-2013***************************
Dim sValor As String
frmListaVinculadosCargo.Show 1
sValor = frmListaVinculadosCargo.psValorSel
Dim existe As Boolean
existe = False
Dim i As Integer
For i = 1 To FEGestion.Rows - 1
    If Mid(sValor, 1, 1) = FEGestion.TextMatrix(i, 1) Then
        existe = True
    End If
Next i
If existe Then
    MsgBox "este valor ya a sido ingresado", vbInformation, "Aviso"
    Call cmdEliminar_Click
    Exit Sub
End If
FEGestion.TextMatrix(FEGestion.row, 1) = Mid(sValor, 1, 1)
FEGestion.TextMatrix(FEGestion.row, 2) = Mid(sValor, 5, Len(sValor))
'JAME FIN ***************************
'FEGestionNoMoverdeFila = FEGestion.Rows - 1
''Call LLenarComboGestion
'FEGestion.lbEditarFlex = True
'FEGestion.SetFocus
End Sub
Public Sub Iniciar(ByVal pnGrupo As Integer, ByVal psPersCodEmpresa As String, ByVal psPersCodVinculado As String)
    lnGrupo = pnGrupo
    lsPersCodEmpresa = psPersCodEmpresa
    lsPersCodVinculado = psPersCodVinculado
    Call MostrarDatos(lnGrupo, lsPersCodEmpresa, lsPersCodVinculado)
    cmdGrabar.Enabled = True
    cmdModificar.Enabled = False
    'cmdEliminar.Enabled = False'JAME 20140303 comentó
    Show 1
End Sub
Private Sub ListarRelacionVinculados(ByVal nGrupo As Integer, ByVal sPersCodEmpresa As String, ByVal sPersCodVinculado As String)
    Dim nRetorno As Integer
    Dim i As Integer
    Dim oGrup As COMDpersona.DCOMGrupoE
    Dim rs As ADODB.Recordset
    With FEGestion
        For i = 1 To .Rows - 1
           FEGestion.EliminaFila (1)
        Next
    End With
    
    Set oGrup = New COMDpersona.DCOMGrupoE
    Set rs = oGrup.ListarPersGrupoEconomicoGestion(nGrupo, sPersCodEmpresa, sPersCodVinculado)
    i = 0
    If rs.EOF Or rs.BOF Then
        nRetorno = 0
    Else
        Do Until rs.EOF
        i = i + 1
            FEGestion.AdicionaFila
            FEGestion.TextMatrix(i, 1) = rs!cRelacionGestion
            FEGestion.TextMatrix(i, 2) = Trim(rs!cDescRelacionGestion)
            rs.MoveNext
        Loop
        nRetorno = 1
    End If
    Set oGrup = Nothing
    rs.Close
End Sub

Private Function MostrarDatos(ByVal nGrupo As Integer, ByVal sPersCodEmpresa As String, ByVal sPersCodVinculado As String) As Integer
    Dim nRetorno As Integer
    Dim oGrup As COMDpersona.DCOMGrupoE
    Dim rs As ADODB.Recordset
    Set oGrup = New COMDpersona.DCOMGrupoE
    Set rs = oGrup.ObtenerPersGrupoEconomico(nGrupo, sPersCodEmpresa, sPersCodVinculado)
    If rs.EOF Or rs.BOF Then
        nRetorno = 0
    Else
        Do Until rs.EOF
            lblGrupoEconomico.Caption = rs!cDesGrupoEconomico
            lblEmpresaVinculado.Caption = rs!cPersEmpresa
            lblPersonaVinculado.Caption = rs!cPersVinculado
            rs.MoveNext
        Loop
        nRetorno = 1
    End If
    If nRetorno = 1 Then
        Call ListarRelacionVinculados(nGrupo, sPersCodEmpresa, sPersCodVinculado)
    End If
    Set oGrup = Nothing
    rs.Close
End Function

Private Sub cmdSalir_Click()
    Unload Me
End Sub

'JAME 20140303 COMENTO
'Private Sub FEGestion_RowColChange()
'If FEGestion.col = 1 Then
'    Dim oGrup As COMDpersona.DCOMGrupoE
'    Set oGrup = New COMDpersona.DCOMGrupoE
'    FEGestion.CargaCombo oGrup.ObtenerRelacionGestion
'    Set oGrup = Nothing
'    If FEGestion.TextMatrix(FEGestion.row, 1) <> "" Then
'        FEGestion.TextMatrix(FEGestion.row, 2) = Mid(FEGestion.TextMatrix(FEGestion.row, 1), 70, Len(FEGestion.TextMatrix(FEGestion.row, 1)) - 70)
'    End If
'End If
'End Sub

