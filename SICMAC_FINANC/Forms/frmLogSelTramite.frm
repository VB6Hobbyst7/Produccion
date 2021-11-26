VERSION 5.00
Object = "{5F774E03-DB36-4DFC-AAC4-D35DC9379F2F}#1.0#0"; "VertMenu.ocx"
Begin VB.Form frmLogSelTramite 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trámite del Proceso de Selección"
   ClientHeight    =   5565
   ClientLeft      =   1080
   ClientTop       =   2025
   ClientWidth     =   10065
   Icon            =   "frmLogSelTramite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   7905
      TabIndex        =   2
      Top             =   5085
      Width           =   1305
   End
   Begin Sicmact.FlexEdit fgeTramite 
      Height          =   4170
      Left            =   1320
      TabIndex        =   0
      Top             =   810
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   7355
      Cols0           =   4
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-Proceso Selección-Resolución-Estado"
      EncabezadosAnchos=   "380-2200-2000-1500"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-L"
      FormatosEdit    =   "0-0-0-0"
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      lbOrdenaCol     =   -1  'True
      ColWidth0       =   375
      RowHeight0      =   285
   End
   Begin Sicmact.Usuario Usuario 
      Left            =   60
      Top             =   5085
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VertMenu.VerticalMenu vmTramite 
      Height          =   4560
      Left            =   135
      TabIndex        =   1
      Top             =   405
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   8043
      MenusMax        =   5
      MenuCur         =   3
      MenuCaption1    =   "Registro"
      MenuItemsMax1   =   4
      MenuItemIcon11  =   "frmLogSelTramite.frx":030A
      MenuItemCaption11=   "Resolución"
      MenuItemIcon12  =   "frmLogSelTramite.frx":0624
      MenuItemCaption12=   "Comité"
      MenuItemIcon13  =   "frmLogSelTramite.frx":093E
      MenuItemCaption13=   "Bases"
      MenuItemIcon14  =   "frmLogSelTramite.frx":0C58
      MenuItemCaption14=   "Parámetros"
      MenuCaption2    =   "Convocatoria"
      MenuItemsMax2   =   2
      MenuItemIcon21  =   "frmLogSelTramite.frx":0F72
      MenuItemCaption21=   "Publicación"
      MenuItemIcon22  =   "frmLogSelTramite.frx":128C
      MenuItemCaption22=   "Cotización"
      MenuCaption3    =   "Bases"
      MenuItemsMax3   =   2
      MenuItemIcon31  =   "frmLogSelTramite.frx":15A6
      MenuItemCaption31=   "Entrega"
      MenuItemIcon32  =   "frmLogSelTramite.frx":18C0
      MenuItemCaption32=   "Observaciones"
      MenuCaption4    =   "Propuesta"
      MenuItemsMax4   =   2
      MenuItemIcon41  =   "frmLogSelTramite.frx":1BDA
      MenuItemCaption41=   "Técnica"
      MenuItemIcon42  =   "frmLogSelTramite.frx":2174
      MenuItemCaption42=   "Económica"
      MenuCaption5    =   "Evaluación"
      MenuItemsMax5   =   2
      MenuItemIcon51  =   "frmLogSelTramite.frx":248E
      MenuItemCaption51=   "Técnica"
      MenuItemIcon52  =   "frmLogSelTramite.frx":27A8
      MenuItemCaption52=   "Económica"
   End
   Begin VB.Label lblAreaDes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1350
      TabIndex        =   5
      Top             =   45
      Width           =   4110
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Area :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   210
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   105
      Width           =   750
   End
   Begin VB.Label lblTitulo 
      BackColor       =   &H8000000C&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Registro"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   1320
      TabIndex        =   3
      Top             =   405
      Width           =   8610
   End
End
Attribute VB_Name = "frmLogSelTramite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim psTpoFrm As String
Dim clsDAdq As DLogAdquisi

Public Sub Inicio(ByVal TipoForm As String)
psTpoFrm = TipoForm
Me.Show 1
End Sub

Private Sub cmdSalir_Click()
    Set clsDAdq = Nothing
    Unload Me
End Sub

Private Sub fgeTramite_DblClick()
    Dim nRow As Integer
    nRow = fgeTramite.Row
    If fgeTramite.TextMatrix(nRow, 1) <> "" Then
        If vmTramite.MenuCur = 1 Then
            'REGISTRO
            If vmTramite.MenuItemCur = 1 Then
                'RESOLUCION
                'Call frmLogReqInicio.Inicio(psTpoFrm, "2", fgeTramite.TextMatrix(nRow, 1))
                Call frmLogSelInicio.Inicio("1")
                Call vmTramite_MenuItemClick(1, 1)
                If fgeTramite.Rows > nRow Then fgeTramite.Row = nRow
            ElseIf vmTramite.MenuItemCur = 2 Then
                'COMITE
                'Call frmLogReqInicio.Inicio(psTpoFrm, "3", fgeTramite.TextMatrix(nRow, 1))
                Call frmLogSelInicio.Inicio("2")
                Call vmTramite_MenuItemClick(1, 2)
                If fgeTramite.Rows > nRow Then fgeTramite.Row = nRow
            ElseIf vmTramite.MenuItemCur = 3 Then
                'BASES
                'Call frmLogReqInicio.Inicio(psTpoFrm, "4", fgeTramite.TextMatrix(nRow, 1))
                Call frmLogSelInicio.Inicio("3")
                Call vmTramite_MenuItemClick(1, 3)
                If fgeTramite.Rows > nRow Then fgeTramite.Row = nRow
            ElseIf vmTramite.MenuItemCur = 4 Then
                'PARAMETROS
                'Call frmLogReqPrecio.Inicio(psTpoFrm, "1", fgeTramite.TextMatrix(nRow, 1))
                Call frmLogSelInicio.Inicio("4")
                Call vmTramite_MenuItemClick(1, 4)
                If fgeTramite.Rows > nRow Then fgeTramite.Row = nRow
            End If
        ElseIf vmTramite.MenuCur = 2 Then
            'CONVOCATORIA
            If vmTramite.MenuItemCur = 1 Then
                'PUBLICACION
                Call frmLogSelInicio.Inicio("5")
                Call vmTramite_MenuItemClick(2, 1)
                If fgeTramite.Rows > nRow Then fgeTramite.Row = nRow
            ElseIf vmTramite.MenuItemCur = 2 Then
                'COTIZACION
                Call frmLogSelCotiza.Inicio("1")
                Call vmTramite_MenuItemClick(2, 2)
                If fgeTramite.Rows > nRow Then fgeTramite.Row = nRow
            End If
        ElseIf vmTramite.MenuCur = 3 Then
            'BASES
            If vmTramite.MenuItemCur = 1 Then
                'ENTREGA
                Call frmLogSelEntBase.Inicio("1")
                Call vmTramite_MenuItemClick(3, 1)
                If fgeTramite.Rows > nRow Then fgeTramite.Row = nRow
            ElseIf vmTramite.MenuItemCur = 2 Then
                'OBSERVACIONES
                Call frmLogSelEntBase.Inicio("2")
                Call vmTramite_MenuItemClick(3, 2)
                If fgeTramite.Rows > nRow Then fgeTramite.Row = nRow
            End If
        ElseIf vmTramite.MenuCur = 4 Then
            'PROPUESTA
            If vmTramite.MenuItemCur = 1 Then
                'TECNICA
                
                Call vmTramite_MenuItemClick(4, 1)
                If fgeTramite.Rows > nRow Then fgeTramite.Row = nRow
            ElseIf vmTramite.MenuItemCur = 2 Then
                'ECONOMICA
                Call frmLogSelCotPro.Inicio("1")
                Call vmTramite_MenuItemClick(4, 2)
                If fgeTramite.Rows > nRow Then fgeTramite.Row = nRow
            End If
        ElseIf vmTramite.MenuCur = 5 Then
            'EVALUACION
            If vmTramite.MenuItemCur = 1 Then
                'TECNICA
                
                Call vmTramite_MenuItemClick(5, 1)
                If fgeTramite.Rows > nRow Then fgeTramite.Row = nRow
            ElseIf vmTramite.MenuItemCur = 2 Then
                'ECONOMICA
                Call frmLogSelCotPro.Inicio("2")
                Call vmTramite_MenuItemClick(5, 2)
                If fgeTramite.Rows > nRow Then fgeTramite.Row = nRow
            End If
        Else
            MsgBox "Botón no reconocida", vbInformation, " Aviso"
        End If
    End If
End Sub

Private Sub Form_Load()
    Set clsDAdq = New DLogAdquisi
    Call CentraForm(Me)
    'Carga información de la relación usuario-area
    Usuario.Inicio gsCodUser
    If Len(Usuario.AreaCod) = 0 Then
        fgeTramite.Enabled = False
        MsgBox "Usuario no determinado", vbInformation, "Aviso"
        Exit Sub
    End If
    lblAreaDes.Caption = Usuario.AreaNom

    If psTpoFrm = "1" Then
        Me.Caption = "Trámite del Proceso de Selección"
    Else
        Me.Caption = ""
    End If
End Sub

Private Sub vmTramite_MenuItemClick(MenuNumber As Long, MenuItem As Long)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    fgeTramite.Clear
    fgeTramite.FormaCabecera
    fgeTramite.Rows = 2
    vmTramite.MenuItemCur = MenuItem
    lblTitulo.Caption = vmTramite.MenuCaption & " : " & vmTramite.MenuItemCaption
    If MenuNumber = 1 Then
        'REGISTRO
        If MenuItem = 1 Then
            'Set rs = clsDAdq.CargaRequerimiento(psTpoFrm, ReqTodosAreaTraNuevo, Usuario.AreaCod)
            Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, "", gLogSelEstadoInicioRes)
            If rs.RecordCount > 0 Then Set fgeTramite.Recordset = rs
        ElseIf MenuItem = 2 Then
            'Set rs = clsDAdq.CargaRequerimiento(psTpoFrm, ReqTodosAreaTraIngreso, Usuario.AreaCod)
            Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, "", gLogSelEstadoInicioRes)
            If rs.RecordCount > 0 Then Set fgeTramite.Recordset = rs
        ElseIf MenuItem = 3 Then
            'Set rs = clsDAdq.CargaRequerimiento(psTpoFrm, ReqTodosAreaTraEgreso, Usuario.AreaCod)
            Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, "", gLogSelEstadoComite)
            If rs.RecordCount > 0 Then Set fgeTramite.Recordset = rs
        ElseIf MenuItem = 4 Then
            'Set rs = clsDAdq.CargaRequerimiento(psTpoFrm, ReqTodosTraPrecio, "")
            Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, "", gLogSelEstadoBases)
            If rs.RecordCount > 0 Then Set fgeTramite.Recordset = rs
        Else
            MsgBox "Opción no definida"
        End If
    ElseIf MenuNumber = 2 Then
        'CONVOCATORIA
        If MenuItem = 1 Then
            Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, "", gLogSelEstadoParametro)
            If rs.RecordCount > 0 Then Set fgeTramite.Recordset = rs
        ElseIf MenuItem = 2 Then
            Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, , gLogSelEstadoPublicacion)
            If rs.RecordCount > 0 Then Set fgeTramite.Recordset = rs
        Else
            MsgBox "Opción en botón 2 no definida"
        End If
    ElseIf MenuNumber = 3 Then
        'BASES
        If MenuItem = 1 Then
            Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, "", gLogSelEstadoRegBase, gLogSelEstadoCotizacion)
            If rs.RecordCount > 0 Then Set fgeTramite.Recordset = rs
        ElseIf MenuItem = 2 Then
            Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, "", gLogSelEstadoRegBase)
            If rs.RecordCount > 0 Then Set fgeTramite.Recordset = rs
        Else
            MsgBox "Opción en botón 3 no definida"
        End If
    ElseIf MenuNumber = 4 Then
        'PROPUESTAS
        If MenuItem = 1 Then
            
        ElseIf MenuItem = 2 Then
            Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, , gLogSelEstadoCotizacion)
            If rs.RecordCount > 0 Then Set fgeTramite.Recordset = rs
        Else
            MsgBox "Opción en botón 4 no definida"
        End If
    ElseIf MenuNumber = 5 Then
        'EVALUACIONES
        If MenuItem = 1 Then
            
        ElseIf MenuItem = 2 Then
            Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, , gLogSelEstadoCotizacion)
            If rs.RecordCount > 0 Then Set fgeTramite.Recordset = rs
        Else
            MsgBox "Opción en botón 5 no definida"
        End If
    Else
        MsgBox "Menú no definido"
    End If
    Set rs = Nothing
End Sub


