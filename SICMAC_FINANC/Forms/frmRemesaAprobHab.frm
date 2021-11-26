VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRemesaAprobHab 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB DE HABILITACIÓN DE REMESAS"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10470
   Icon            =   "frmRemesaAprobHab.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraBusqueda 
      Height          =   1095
      Left            =   80
      TabIndex        =   6
      Top             =   0
      Width           =   10335
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
         Height          =   320
         Left            =   9120
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.Frame fraFecha 
         Caption         =   "Fecha"
         Enabled         =   0   'False
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
         Height          =   855
         Left            =   165
         TabIndex        =   7
         Top             =   120
         Width           =   1410
         Begin MSMask.MaskEdBox txtFecha 
            Height          =   315
            Left            =   120
            TabIndex        =   8
            Top             =   330
            Width           =   1155
            _ExtentX        =   2037
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
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   320
      Left            =   9320
      TabIndex        =   5
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdAprobar 
      Caption         =   "&Aprobar"
      Height          =   320
      Left            =   80
      TabIndex        =   4
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdRechazar 
      Caption         =   "&Rechazar"
      Height          =   320
      Left            =   1175
      TabIndex        =   3
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   320
      Left            =   8205
      TabIndex        =   2
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CheckBox chkTodos 
      Appearance      =   0  'Flat
      Caption         =   "Todos"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   195
      TabIndex        =   0
      Top             =   1150
      Width           =   780
   End
   Begin Sicmact.FlexEdit fg 
      Height          =   3885
      Left            =   75
      TabIndex        =   1
      Top             =   1440
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6853
      Cols0           =   8
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Itm-Fecha-Agencia Origen-Agencia Destino-Moneda-Importe-ID"
      EncabezadosAnchos=   "400-500-1750-2500-2500-1000-1250-0"
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-1-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-4-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-L-L-C-R-L"
      FormatosEdit    =   "0-0-0-0-0-0-2-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbPuntero       =   -1  'True
      lbOrdenaCol     =   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmRemesaAprobHab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************
'** Nombre : frmRemesaAprobHab
'** Descripción : Formulario para dar VB de las solicitudes de Habilitación de Remesas
'** Creación : EJVG, 20140711 15:30:00 PM
'*************************************************************************************
Option Explicit
Dim fsOpeCod As String

Private Sub chkTodos_Click()
    Dim i As Long
    Dim lsCheck As String
    If fg.TextMatrix(1, 0) = "" Then
        chkTodos.value = 0
        Exit Sub
    End If
    lsCheck = IIf(chkTodos.value = 1, "1", "")
    For i = 1 To fg.Rows - 1
        fg.TextMatrix(i, 1) = lsCheck
    Next
End Sub
Private Sub cmdCancelar_Click()
    Limpiar
    fraBusqueda.Enabled = True
End Sub
Private Sub cmdProcesar_Click()
    Dim oCaja As nCajaGeneral
    Dim rs As ADODB.Recordset
    Dim i As Long
    Dim lsMarca As String
    
    On Error GoTo ErrcmdProcesar
    If Not ValidaInterfaz Then Exit Sub
    
    Set oCaja = New nCajaGeneral
    Set rs = New ADODB.Recordset
    chkTodos.value = 0
    FormateaFlex fg
    
    Screen.MousePointer = 11
    Set rs = oCaja.ListaSolicitudHabRemesaxAprobacion(Mid(fsOpeCod, 3, 1), CDate(txtFecha.Text))
    If Not rs.EOF Then
        lsMarca = "1"
        Do While Not rs.EOF
            fg.AdicionaFila
            i = fg.row
            fg.TextMatrix(i, 1) = lsMarca
            fg.TextMatrix(i, 2) = Format(rs!dFecha, "dd/mm/yyyy hh:mm:ss AMPM")
            fg.TextMatrix(i, 3) = rs!cAgeOrigen
            fg.TextMatrix(i, 4) = rs!cAgeDestino
            fg.TextMatrix(i, 5) = rs!cMoneda
            fg.TextMatrix(i, 6) = Format(rs!nImporte, gsFormatoNumeroView)
            fg.TextMatrix(i, 7) = rs!nId
            rs.MoveNext
        Loop
        fraBusqueda.Enabled = False
    Else
        lsMarca = "0"
        MsgBox "No se encontraron resultados", vbInformation, "Aviso"
    End If
    chkTodos.value = lsMarca
    RSClose rs
    Set oCaja = Nothing
    Screen.MousePointer = 0
    Exit Sub
ErrcmdProcesar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Public Sub Inicia(ByVal psOpeCod As String, ByVal psOpeDesc As String)
    fsOpeCod = psOpeCod
    Caption = UCase(psOpeDesc)
    Show 1
End Sub
Private Sub Form_Load()
    Limpiar
End Sub
Private Sub Limpiar()
    txtFecha.Text = Format(gdFecSis, gsFormatoFechaView)
    chkTodos.value = 0
    FormateaFlex fg
End Sub
Private Function ValidaInterfaz() As Boolean
    Dim lsValFecha As String
    ValidaInterfaz = False
    lsValFecha = ValidaFecha(txtFecha)
    If Len(lsValFecha) > 0 Then
        MsgBox lsValFecha, vbInformation, "Aviso"
        If fraFecha.Visible And fraFecha.Enabled Then EnfocaControl txtFecha
        Exit Function
    End If
    ValidaInterfaz = True
End Function
Private Sub cmdAprobar_Click()
    Dim oCaja As nCajaGeneral
    Dim oContFunciones As NContFunciones
    Dim lsIDs As String, lsMovNro As String
    Dim lbExito As Boolean
    
    On Error GoTo ErrCmdAprobar
    lsIDs = DameListaIDs
    If Len(lsIDs) = 0 Then
        MsgBox "Ud. debe seleccionar por lo menos una solicitud para continuar", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("¿Esta seguro de aprobar las solicitudes seleccionadas?", vbYesNo + vbQuestion, "Aviso") = vbNo Then Exit Sub
    Set oCaja = New nCajaGeneral
    Set oContFunciones = New NContFunciones
    lsMovNro = oContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    
    lbExito = oCaja.AprobarSolicitudHabRemesa(lsIDs, True, lsMovNro)
    If lbExito Then
        MsgBox "Se ha aprobado satisfactoriamente las solicitudes seleccionadas", vbInformation, "Aviso"
        cmdCancelar_Click
    Else
        MsgBox "Ha sucedido un error al aprobar las solicitudes seleccionadas, si esto persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If
    Set oCaja = Nothing
    Set oContFunciones = Nothing
    Exit Sub
ErrCmdAprobar:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdRechazar_Click()
    Dim oCaja As nCajaGeneral
    Dim oContFunciones As NContFunciones
    Dim lsIDs As String, lsMovNro As String
    Dim lbExito As Boolean
    
    On Error GoTo ErrCmdRechazar
    lsIDs = DameListaIDs
    If Len(lsIDs) = 0 Then
        MsgBox "Ud. debe seleccionar por lo menos una solicitud para continuar", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("¿Esta seguro de rechazar las solicitudes seleccionadas?", vbYesNo + vbQuestion, "Aviso") = vbNo Then Exit Sub
    Set oCaja = New nCajaGeneral
    Set oContFunciones = New NContFunciones
    lsMovNro = oContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    
    lbExito = oCaja.AprobarSolicitudHabRemesa(lsIDs, False, lsMovNro)
    If lbExito Then
        MsgBox "Se ha rechazado satisfactoriamente las solicitudes seleccionadas", vbInformation, "Aviso"
        cmdCancelar_Click
    Else
        MsgBox "Ha sucedido un error al rechazar las solicitudes seleccionadas, si esto persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If
    Set oCaja = Nothing
    Set oContFunciones = Nothing
    Exit Sub
ErrCmdRechazar:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Function DameListaIDs() As String
    Dim fila As Long
    Dim Lista As String
    If fg.TextMatrix(1, 0) <> "" Then
        For fila = 1 To fg.Rows - 1
            If fg.TextMatrix(fila, 1) = "." Then
                Lista = Lista & fg.TextMatrix(fila, 7) & ","
            End If
        Next
        If Len(Lista) > 0 Then
            Lista = Mid(Lista, 1, Len(Lista) - 1)
        End If
    End If
    DameListaIDs = Lista
End Function
