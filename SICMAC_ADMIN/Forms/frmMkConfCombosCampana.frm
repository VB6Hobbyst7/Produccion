VERSION 5.00
Begin VB.Form frmMkConfCombosCampana 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de Combos por Campaña"
   ClientHeight    =   5550
   ClientLeft      =   9585
   ClientTop       =   4695
   ClientWidth     =   7230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMkConfCombosCampana.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7230
   Begin Sicmact.FlexEdit flxCampanas 
      Height          =   3375
      Left            =   360
      TabIndex        =   7
      Top             =   1200
      Width           =   6255
      _extentx        =   11033
      _extenty        =   6800
      cols0           =   12
      highlight       =   1
      allowuserresizing=   3
      rowsizingmode   =   1
      encabezadosnombres=   "#-nIdCombo-nIdCampana-Descripcion-bDesembolso-bApertura-bSoles-bDolares-nMinSoles-nMaxSoles-nMinDolares-nMaxDolares"
      encabezadosanchos=   "500-0-0-4200-0-0-0-0-0-0-0-0"
      font            =   "frmMkConfCombosCampana.frx":030A
      font            =   "frmMkConfCombosCampana.frx":0336
      font            =   "frmMkConfCombosCampana.frx":0362
      font            =   "frmMkConfCombosCampana.frx":038E
      font            =   "frmMkConfCombosCampana.frx":03BA
      fontfixed       =   "frmMkConfCombosCampana.frx":03E6
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1
      columnasaeditar =   "X-X-X-X-X-X-X-X-X-X-X-X"
      textstylefixed  =   3
      listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
      encabezadosalineacion=   "C-C-C-C-C-C-C-C-C-C-C-C"
      formatosedit    =   "0-0-0-0-0-0-0-0-0-0-0-0"
      textarray0      =   "#"
      lbbuscaduplicadotext=   -1
      colwidth0       =   495
      rowheight0      =   300
      forecolorfixed  =   -2147483630
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   360
      Left            =   5520
      TabIndex        =   6
      Top             =   4800
      Width           =   990
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "Quitar"
      Height          =   360
      Left            =   3600
      TabIndex        =   5
      Top             =   4800
      Width           =   990
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "Ver"
      Height          =   360
      Left            =   2520
      TabIndex        =   4
      Top             =   4800
      Width           =   990
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   360
      Left            =   1440
      TabIndex        =   3
      Top             =   4800
      Width           =   990
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      Height          =   360
      Left            =   360
      TabIndex        =   2
      Top             =   4800
      Width           =   990
   End
   Begin VB.ComboBox cmbCampaña 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   5415
   End
   Begin VB.Label lblCombosPor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Combos por Campaña:"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   840
      Width           =   1635
   End
   Begin VB.Label lblCampaña 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Campaña:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frmMkConfCombosCampana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oNGastosMarketing As New NGastosMarketing
Const msgNohayDatos = "No hay datos"
Private Type itemComboBox
    cod As String
    dsc As String
End Type
Dim lsCampanas() As itemComboBox

Private Sub cmbCampaña_Click()
    Dim idCampana As String
    idCampana = getIdLista(cmbCampaña.ListIndex, lsCampanas)
    llenarCombosPorCampana idCampana
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdQuitar_Click()
    
    If flxCampanas.TextMatrix(flxCampanas.row, 3) <> "" Then
        
        Dim pregunta As String
        'Preguntamos si esta seguro de grabar
        pregunta = MsgBox("¿Está Seguro que va a quitar el combo?.", vbYesNo + vbExclamation + vbDefaultButton2, "Eliminar Combo.")
        If pregunta <> vbYes Then
            Exit Sub
        End If
        
        Dim idCampana As String
        idCampana = getIdLista(cmbCampaña.ListIndex, lsCampanas)
        Dim idcombo As String: idcombo = flxCampanas.TextMatrix(flxCampanas.row, 1)
        oNGastosMarketing.EliminaComboCampana idcombo
        llenarCombosPorCampana idCampana
    End If
    
End Sub

Private Sub cmdVer_Click()
    If flxCampanas.TextMatrix(flxCampanas.row, 3) <> "" Then
        Dim idCampana As String
        Dim idcombo As String: idcombo = flxCampanas.TextMatrix(flxCampanas.row, 1)
        Dim cComboDescripcion As String: cComboDescripcion = flxCampanas.TextMatrix(flxCampanas.row, 3)
        Dim bDesembolso As Integer: bDesembolso = IIf((flxCampanas.TextMatrix(flxCampanas.row, 4)) = "Falso", 0, 1)
        Dim bApertura As Integer: bApertura = IIf((flxCampanas.TextMatrix(flxCampanas.row, 5)) = "Falso", 0, 1)
        Dim bSoles As Integer: bSoles = IIf((flxCampanas.TextMatrix(flxCampanas.row, 6)) = "Falso", 0, 1)
        Dim bDolares As Integer: bDolares = IIf((flxCampanas.TextMatrix(flxCampanas.row, 7)) = "Falso", 0, 1)
        Dim nMinSoles As String: nMinSoles = (flxCampanas.TextMatrix(flxCampanas.row, 8))
        Dim nMaxSoles As String: nMaxSoles = (flxCampanas.TextMatrix(flxCampanas.row, 9))
        Dim nMinDolares As String: nMinDolares = flxCampanas.TextMatrix(flxCampanas.row, 10)
        Dim nMaxDolares As String: nMaxDolares = flxCampanas.TextMatrix(flxCampanas.row, 11)
        idCampana = getIdLista(cmbCampaña.ListIndex, lsCampanas)
'        If idCampana <> "0" Then
            frmMkNuevoCombo.aDetalle idcombo, idCampana, cComboDescripcion, bDesembolso, bApertura, bSoles, bDolares, nMinSoles, nMaxSoles, nMinDolares, nMaxDolares
'        Else
'            MsgBox "Debe elegir una campaña"
'        End If
    End If
End Sub

Private Sub cmdEditar_Click()
    If flxCampanas.TextMatrix(flxCampanas.row, 3) <> "" Then
        Dim idCampana As String
        Dim idcombo As String: idcombo = flxCampanas.TextMatrix(flxCampanas.row, 1)
        Dim cComboDescripcion As String: cComboDescripcion = flxCampanas.TextMatrix(flxCampanas.row, 3)
        Dim bDesembolso As Integer: bDesembolso = IIf((flxCampanas.TextMatrix(flxCampanas.row, 4)) = "Falso", 0, 1)
        Dim bApertura As Integer: bApertura = IIf((flxCampanas.TextMatrix(flxCampanas.row, 5)) = "Falso", 0, 1)
        Dim bSoles As Integer: bSoles = IIf((flxCampanas.TextMatrix(flxCampanas.row, 6)) = "Falso", 0, 1)
        Dim bDolares As Integer: bDolares = IIf((flxCampanas.TextMatrix(flxCampanas.row, 7)) = "Falso", 0, 1)
        Dim nMinSoles As String: nMinSoles = (flxCampanas.TextMatrix(flxCampanas.row, 8))
        Dim nMaxSoles As String: nMaxSoles = (flxCampanas.TextMatrix(flxCampanas.row, 9))
        Dim nMinDolares As String: nMinDolares = flxCampanas.TextMatrix(flxCampanas.row, 10)
        Dim nMaxDolares As String: nMaxDolares = flxCampanas.TextMatrix(flxCampanas.row, 11)
        idCampana = getIdLista(cmbCampaña.ListIndex, lsCampanas)
'        If idCampana <> "0" Then
            frmMkNuevoCombo.aEditar idcombo, idCampana, cComboDescripcion, bDesembolso, bApertura, bSoles, bDolares, nMinSoles, nMaxSoles, nMinDolares, nMaxDolares, "Editar Combo de Campaña"
'        Else
'            MsgBox "Debe elegir una campaña"
'        End If
    llenarCombosPorCampana idCampana
    End If
End Sub

Private Sub cmdNuevo_Click()
    Dim idCampana As String
    idCampana = getIdLista(cmbCampaña.ListIndex, lsCampanas)
    
'    If idCampana <> "0" Then
        frmMkNuevoCombo.aNuevo idCampana
        llenarCombosPorCampana idCampana
'    Else
'        MsgBox "Debe elegir una campaña"
'    End If
End Sub

Private Sub Form_Load()
    Call llenarComboCampanasActivas
End Sub
Private Sub llenarComboCampanasActivas()
    Dim rs As ADODB.Recordset
    Set rs = oNGastosMarketing.RecuperaCampanas
    Dim n As Integer
    n = 0
    Do While Not rs.EOF
        ReDim Preserve lsCampanas(n)
        lsCampanas(n).cod = rs!nConsValor
        lsCampanas(n).dsc = rs!cConsDescripcion
        cmbCampaña.AddItem Trim(lsCampanas(n).dsc)
        n = n + 1
        rs.MoveNext
    Loop
    cmbCampaña = cmbCampaña.List(0)
End Sub
Private Sub llenarCombosPorCampana(ByVal idCampana As String)
    Dim rs As ADODB.Recordset
    Set rs = oNGastosMarketing.RecuperaCombosXCampana(idCampana)
    Dim n As Integer
    n = 0
    'flxCampanas.Rows = 2
    'flxCampanas.Clear
    Call LimpiaFlex(flxCampanas)
    flxCampanas.TextMatrix(1, 3) = msgNohayDatos
    Do While Not rs.EOF
        flxCampanas.AdicionaFila
        flxCampanas.TextMatrix(n + 1, 1) = rs!nIdCombo
        flxCampanas.TextMatrix(n + 1, 2) = rs!nIdCampana
        flxCampanas.TextMatrix(n + 1, 3) = rs!cComboDescripcion
        flxCampanas.TextMatrix(n + 1, 4) = IIf(IsNull(rs!bDesembolso), "Falso", rs!bDesembolso)
        flxCampanas.TextMatrix(n + 1, 5) = IIf(IsNull(rs!bApertura), "Falso", rs!bApertura)
        flxCampanas.TextMatrix(n + 1, 6) = IIf(IsNull(rs!bSoles), "Falso", rs!bSoles)
        flxCampanas.TextMatrix(n + 1, 7) = IIf(IsNull(rs!bDolares), "Falso", rs!bDolares)
        flxCampanas.TextMatrix(n + 1, 8) = IIf(IsNull(rs!nMinSoles), "", rs!nMinSoles)
        flxCampanas.TextMatrix(n + 1, 9) = IIf(IsNull(rs!nMaxSoles), "", rs!nMaxSoles)
        flxCampanas.TextMatrix(n + 1, 10) = IIf(IsNull(rs!nMinDolares), "", rs!nMinDolares)
        flxCampanas.TextMatrix(n + 1, 11) = IIf(IsNull(rs!nMaxDolares), "", rs!nMaxDolares)
        n = n + 1
        rs.MoveNext
    Loop
    
End Sub
Private Function getIdLista(ByVal index As Integer, ByRef item() As itemComboBox) As String
    getIdLista = item(index).cod
End Function

Private Sub cmdAceptar_Click()
    'aqui va la logica de registro
End Sub

