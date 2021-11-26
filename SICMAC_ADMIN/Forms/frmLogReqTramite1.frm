VERSION 5.00
Object = "{5F774E03-DB36-4DFC-AAC4-D35DC9379F2F}#1.1#0"; "VertMenu.ocx"
Begin VB.Form frmLogReqTramite 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trámite de la Proyección de Requerimiento"
   ClientHeight    =   5565
   ClientLeft      =   450
   ClientTop       =   2040
   ClientWidth     =   11190
   Icon            =   "frmLogReqTramite1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   Begin VertMenu.VerticalMenu vmTramite 
      Height          =   4170
      Left            =   60
      TabIndex        =   5
      Top             =   825
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   7355
      MenuCaption1    =   "Menu1"
      MenuItemIcon11  =   "frmLogReqTramite1.frx":030A
   End
   Begin Sicmact.FlexEdit fgeTramite 
      Height          =   4170
      Left            =   1290
      TabIndex        =   1
      Top             =   810
      Width           =   9735
      _extentx        =   17171
      _extenty        =   7355
      cols0           =   7
      highlight       =   1
      allowuserresizing=   3
      encabezadosnombres=   "Item-Nro.Requerimiento-Procedencia-Periodo-Necesidad-Requerimiento-Estado"
      encabezadosanchos=   "380-2100-1800-600-1800-1800-800"
      font            =   "frmLogReqTramite1.frx":0624
      font            =   "frmLogReqTramite1.frx":0658
      font            =   "frmLogReqTramite1.frx":068C
      font            =   "frmLogReqTramite1.frx":06C0
      font            =   "frmLogReqTramite1.frx":06F4
      fontfixed       =   "frmLogReqTramite1.frx":0728
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1
      columnasaeditar =   "X-X-X-X-X-X-X"
      textstylefixed  =   3
      listacontroles  =   "0-0-0-0-0-0-0"
      encabezadosalineacion=   "C-L-L-C-L-L-L"
      formatosedit    =   "0-0-0-0-0-0-0"
      textarray0      =   "Item"
      lbordenacol     =   -1
      colwidth0       =   375
      rowheight0      =   285
   End
   Begin Sicmact.Usuario Usuario 
      Left            =   30
      Top             =   5085
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   8760
      TabIndex        =   0
      Top             =   5070
      Width           =   1305
   End
   Begin VB.Label lblTitulo 
      BackColor       =   &H8000000C&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1290
      TabIndex        =   4
      Top             =   405
      Width           =   9720
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
      Left            =   450
      TabIndex        =   3
      Top             =   105
      Width           =   750
   End
   Begin VB.Label lblAreaDes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1320
      TabIndex        =   2
      Top             =   45
      Width           =   4110
   End
End
Attribute VB_Name = "frmLogReqTramite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim psTpoReq As String
Dim clsDReq As DLogRequeri

Public Sub Inicio(ByVal TipoReq As String)
psTpoReq = TipoReq
If psTpoReq = "1" Then
    Me.Caption = "Requerimiento Regular : Trámite "
Else
    Me.Caption = "Requerimiento Extemporaneo : Trámite "
End If
Me.Show 1
End Sub

Private Sub cmdSalir_Click()
    Set clsDReq = Nothing
    Unload Me
End Sub

Private Sub fgeTramite_DblClick()
    Dim nRow As Integer
    nRow = fgeTramite.Row
    If fgeTramite.TextMatrix(nRow, 1) <> "" Then
        If vmTramite.MenuCur = 1 Then
            If vmTramite.MenuItemCur = 1 Then
                'NUEVOS
                Call frmLogReqInicio.Inicio(psTpoReq, "2", fgeTramite.TextMatrix(nRow, 1))
                Call vmTramite_MenuItemClick(1, 1)
                If fgeTramite.Rows > nRow Then fgeTramite.Row = nRow
            ElseIf vmTramite.MenuItemCur = 2 Then
                'TRAMITADOS
                Call frmLogReqInicio.Inicio(psTpoReq, "4", fgeTramite.TextMatrix(nRow, 1))
                Call vmTramite_MenuItemClick(1, 2)
                If fgeTramite.Rows > nRow Then fgeTramite.Row = nRow
            'ElseIf vmTramite.MenuItemCur = 4 Then
            '    If Usuario.AreaTrami = gLogAreaTraEstadoPrecio Then
            '        'PRECIOS
            '        Call frmLogReqPrecio.Inicio(psTpoReq, "1", fgeTramite.TextMatrix(nRow, 1))
            '    End If
            '
            '    Call vmTramite_MenuItemClick(1, 4)
            '    If fgeTramite.Rows > nRow Then fgeTramite.Row = nRow
            End If
        ElseIf vmTramite.MenuCur = 2 Then
            If vmTramite.MenuItemCur = 1 Then
                'RECEPCIONADOS
                Call frmLogReqInicio.Inicio(psTpoReq, "3", fgeTramite.TextMatrix(nRow, 1))
                Call vmTramite_MenuItemClick(2, 1)
                If fgeTramite.Rows > nRow Then fgeTramite.Row = nRow
            ElseIf vmTramite.MenuItemCur = 2 Then
                'TRAMITADOS
                Call frmLogReqInicio.Inicio(psTpoReq, "4", fgeTramite.TextMatrix(nRow, 1))
                Call vmTramite_MenuItemClick(2, 2)
                If fgeTramite.Rows > nRow Then fgeTramite.Row = nRow
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Set clsDReq = New DLogRequeri
    Call CentraForm(Me)
    'Carga información de la relación usuario-area
    Usuario.Inicio gsCodUser
    If Len(Usuario.AreaCod) = 0 Then
        fgeTramite.Enabled = False
        MsgBox "Usuario no determinado", vbInformation, "Aviso"
        Exit Sub
    End If
    lblAreaDes.Caption = Usuario.AreaNom
End Sub

Private Sub vmTramite_MenuItemClick(MenuNumber As Long, MenuItem As Long)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    fgeTramite.Clear
    fgeTramite.FormaCabecera
    fgeTramite.Rows = 2
    vmTramite.MenuCur = MenuNumber
    vmTramite.MenuItemCur = MenuItem
    lblTitulo.Caption = vmTramite.MenuCaption & " : " & vmTramite.MenuItemCaption
    If MenuNumber = 1 Then
        'PROPIOS
        If MenuItem = 1 Then
            'NUEVOS
            Set rs = clsDReq.CargaRequerimiento(psTpoReq, ReqTodosAreaTraNuevo, Usuario.AreaCod)
            If rs.RecordCount > 0 Then Set fgeTramite.Recordset = rs
        ElseIf MenuItem = 2 Then
            'TRAMITADOS
            Set rs = clsDReq.CargaRequerimiento(psTpoReq, ReqTodosAreaTraEgreso, Usuario.AreaCod)
            If rs.RecordCount > 0 Then Set fgeTramite.Recordset = rs
        'ElseIf MenuItem = 4 Then
        '    If Usuario.AreaTrami = gLogAreaTraEstadoPrecio Then
        '        Set rs = clsDReq.CargaRequerimiento(psTpoReq, ReqTodosTraPrecio, "")
        '        If rs.RecordCount > 0 Then Set fgeTramite.Recordset = rs
        '    Else
        '        MsgBox "Opción en botón 4 no definida"
        '    End If
        Else
            MsgBox "Opción no definida"
        End If
    ElseIf MenuNumber = 2 Then
        'OTROS
        If MenuItem = 1 Then
            'RECEPCIONADOS
            Set rs = clsDReq.CargaRequerimiento(psTpoReq, ReqTodosAreaTraIngreso, Usuario.AreaCod)
            If rs.RecordCount > 0 Then Set fgeTramite.Recordset = rs
        ElseIf MenuItem = 2 Then
            'TRAMITADOS
            Set rs = clsDReq.CargaRequerimiento(psTpoReq, ReqTodosAreaTraEgresoOtro, Usuario.AreaCod)
            If rs.RecordCount > 0 Then Set fgeTramite.Recordset = rs
        Else
            MsgBox "Opción no definida"
        End If
    Else
            MsgBox "Menú no definido"
    End If
    Set rs = Nothing
End Sub

