VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRemesaConfirmacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONFIRMACIÓN DE REMESAS RECIBIDAS"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8820
   Icon            =   "frmRemesaConfirmacion.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5460
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   8655
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   320
         Left            =   6330
         TabIndex        =   16
         Top             =   5025
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "&Confirmar"
         Height          =   320
         Left            =   120
         TabIndex        =   15
         Top             =   5025
         Width           =   1095
      End
      Begin VB.TextBox txtGlosa 
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   4320
         Width           =   8415
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         Caption         =   "Todos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   780
      End
      Begin VB.CommandButton cmdRechazar 
         Caption         =   "&Rechazar"
         Height          =   320
         Left            =   1225
         TabIndex        =   10
         Top             =   5025
         Width           =   1095
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   320
         Left            =   7440
         TabIndex        =   9
         Top             =   5025
         Width           =   1095
      End
      Begin SICMACT.FlexEdit fg 
         Height          =   3525
         Left            =   120
         TabIndex        =   12
         Top             =   525
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   6218
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Itm-Fecha-Origen-Moneda-Monto Hab.-Tipo Transporte-Empresa-nMovNro-cMovNro"
         EncabezadosAnchos=   "400-500-1750-2500-1000-1250-1500-1800-0-0"
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
         ColumnasAEditar =   "X-1-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-4-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-L-C-R-L-L-C-C"
         FormatosEdit    =   "0-0-0-0-0-2-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label1 
         Caption         =   "Glosa :"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   4080
         Width           =   735
      End
   End
   Begin VB.Frame fraBusqueda 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
         Height          =   320
         Left            =   7440
         TabIndex        =   7
         Top             =   480
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
         Height          =   735
         Left            =   5400
         TabIndex        =   5
         Top             =   240
         Width           =   1575
         Begin MSMask.MaskEdBox txtFecha 
            Height          =   330
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
      End
      Begin VB.Frame fraOrigen 
         Caption         =   "Destino"
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
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5175
         Begin SICMACT.TxtBuscar txtAreaAgeCod 
            Height          =   300
            Left            =   915
            TabIndex        =   2
            Top             =   240
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            sTitulo         =   ""
         End
         Begin VB.Label Label6 
            Caption         =   "Agencia :"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   270
            Width           =   735
         End
         Begin VB.Label lblAreaAgeDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2085
            TabIndex        =   3
            Top             =   240
            Width           =   2955
         End
      End
   End
End
Attribute VB_Name = "frmRemesaConfirmacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'** Nombre : frmRemesaConfirmacion
'** Descripción : Formulario para el remesas de Agencias a Agencias o a Inst. Financieras
'** Creación : EJVG, 20140630 11:00:00 AM
'****************************************************************************************
Option Explicit

Dim fnMoneda As Moneda
Dim fsCtaContCodH As String
Dim fsCtaContCodD As String

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
    Dim oCaja As COMNCajaGeneral.NCOMCajaGeneral
    Dim rs As New ADODB.Recordset
    Dim fila As Long
    Dim lsMarca As String
    
    On Error GoTo ErrcmdProcesar
    If Not ValidaInterfaz Then Exit Sub
    
    Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
    Set rs = New ADODB.Recordset
    chkTodos.value = 0
    FormateaFlex fg
    
    Screen.MousePointer = 11
    Set rs = oCaja.ListaHabilitacionRemesa(Right(txtAreaAgeCod, 2), CDate(txtFecha.Text))
    If Not rs.EOF Then
        lsMarca = "1"
        Do While Not rs.EOF
            fg.AdicionaFila
            fila = fg.row
            fg.TextMatrix(fila, 1) = lsMarca
            fg.TextMatrix(fila, 2) = Format(rs!dFecha, "dd/mm/yyyy hh:mm:ss AMPM")
            fg.TextMatrix(fila, 3) = rs!cAgeOrigen
            fg.TextMatrix(fila, 4) = rs!cMoneda
            fg.TextMatrix(fila, 5) = Format(rs!nMovImporte, gsFormatoNumeroView)
            fg.TextMatrix(fila, 6) = rs!cTipoTransp
            fg.TextMatrix(fila, 7) = rs!cPersNombreTransp
            fg.TextMatrix(fila, 8) = rs!nMovNro
            fg.TextMatrix(fila, 9) = rs!cMovNro
            rs.MoveNext
        Loop
        'fraBusqueda.Enabled = False
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
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    CargarControles
    Limpiar
End Sub
Private Sub CargarControles()
    Dim oNCont As New NConstSistemas

    On Error GoTo ErrCargarControles
    Screen.MousePointer = 11
    'Carga Cuentas Contables
    fsCtaContCodH = oNCont.LeeConstSistema(476)
    fsCtaContCodD = oNCont.LeeConstSistema(477)
    
    Set oNCont = Nothing
    Screen.MousePointer = 0
    Exit Sub
ErrCargarControles:
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
Private Sub Limpiar()
    txtAreaAgeCod.Text = "026" & Right(gsCodAge, 2)
    lblAreaAgeDesc.Caption = gsNomAge
    txtFecha.Text = Format(gdFecSis, gsFormatoFechaView)
    chkTodos.value = 0
    FormateaFlex fg
    txtGlosa.Text = ""
End Sub
Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        EnfocaControl cmdConfirmar
    End If
End Sub
Private Function ValidaInterfaz() As Boolean
    Dim lsValFecha As String
    ValidaInterfaz = False
    If Len(txtAreaAgeCod.Text) <> 5 Then
        MsgBox "No se ha especificado la Agencia Destino", vbInformation, "Aviso"
        Exit Function
    End If
    lsValFecha = ValidaFecha(txtFecha.Text)
    If Len(lsValFecha) > 0 Then
        MsgBox lsValFecha, vbInformation, "Aviso"
        Exit Function
    End If
    ValidaInterfaz = True
End Function
Private Function DameListaMovimientos(ByVal psAgeCod As String) As Variant
    Dim fila As Long
    Dim Lista As Variant
    Dim iLista As Integer
    Dim lsCtaContTemp As String
    Dim oNContFunciones As New NContFunciones
    
    ReDim Lista(1 To 9, 0 To 0)
    If fg.TextMatrix(1, 0) <> "" Then
        For fila = 1 To fg.Rows - 1
            If fg.TextMatrix(fila, 1) = "." Then
                iLista = UBound(Lista, 2) + 1
                ReDim Preserve Lista(1 To 9, 0 To iLista)
                Lista(1, iLista) = CLng(fg.TextMatrix(fila, 8)) 'nMovNroRef
                Lista(2, iLista) = IIf(UCase(fg.TextMatrix(fila, 4)) = "SOLES", gMonedaNacional, gMonedaExtranjera) 'Moneda
                Lista(3, iLista) = CCur(fg.TextMatrix(fila, 5)) 'Monto
                Lista(4, iLista) = 0 'nMovNro de regreso
                Lista(5, iLista) = "" 'sMovNro de regreso
                Lista(6, iLista) = ReemplazaCaracterCtaCont(fsCtaContCodD, Lista(2, iLista), psAgeCod) 'CtaCont Debe
                'Verifica cuenta de caja de Agencia destino si tiene puente
                lsCtaContTemp = oNContFunciones.BuscaCtaEquivalente(Lista(6, iLista))
                If Len(lsCtaContTemp) > 0 Then Lista(6, iLista) = lsCtaContTemp
                Lista(7, iLista) = ReemplazaCaracterCtaCont(fsCtaContCodH, Lista(2, iLista), psAgeCod) 'CtaCont Haber
                Lista(8, iLista) = fg.TextMatrix(fila, 9) 'cMovNroRef
                Lista(9, iLista) = fg.TextMatrix(fila, 0) 'Item del Flex
            End If
        Next
    End If
    Set oNContFunciones = Nothing
    DameListaMovimientos = Lista
End Function
Private Function ReemplazaCaracterCtaCont(ByVal psCtaContCod As String, ByVal pnMoneda As Moneda, ByVal psAgeCod As String) As String
    ReemplazaCaracterCtaCont = psCtaContCod
    ReemplazaCaracterCtaCont = Replace(ReemplazaCaracterCtaCont, "M", pnMoneda)
    ReemplazaCaracterCtaCont = Replace(ReemplazaCaracterCtaCont, "AG", Format(psAgeCod, "00"))
End Function
Private Sub cmdConfirmar_Click()
    Dim oCaja As COMNCajaGeneral.NCOMCajaGeneral
    Dim oNContFunciones As clases.NContFunciones
    Dim oImpre As clases.NContImprimir
    Dim oPrevio As clsprevio
    Dim Datos As Variant
    Dim lsAgeCod As String
    Dim i As Integer
    Dim lbExito As Boolean
    Dim lsCadImpre As String
    
    On Error GoTo ErrCmdConfirmar
    
    lsAgeCod = Right(txtAreaAgeCod.Text, 2)
    Datos = DameListaMovimientos(lsAgeCod)
    If UBound(Datos, 2) = 0 Then
        MsgBox "Ud. debe seleccionar al menos una Habilitación para continuar", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(Trim(txtGlosa.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la glosa de la confirmación", vbInformation, "Aviso"
        EnfocaControl txtGlosa
        Exit Sub
    End If
    'Valida cuentas contables a ultimo nivel
    'Set oNContFunciones = New clases.NContFunciones
    'For i = 1 To UBound(Datos, 2)
    '    If Not oNContFunciones.verificarUltimoNivelCta(Datos(6, i)) Then
    '        MsgBox "La cuenta contable " & Datos(6, i) & " no es de Ultimo Nivel, comunicarse con el Dpto. de Contabilidad", vbInformation, "Aviso"
    '        Set oNContFunciones = Nothing
    '        Exit Sub
    '    End If
    '    If Not oNContFunciones.verificarUltimoNivelCta(Datos(7, i)) Then
    '        MsgBox "La cuenta contable " & Datos(7, i) & " no es de Ultimo Nivel, comunicarse con el Dpto. de Contabilidad", vbInformation, "Aviso"
    '        Set oNContFunciones = Nothing
    '        Exit Sub
    '    End If
    'Next
    
    If MsgBox("¿Esta seguro de confirmar la Habilitación de la remesa?", vbYesNo + vbQuestion, "Aviso") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    cmdConfirmar.Enabled = False
    Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
    lbExito = oCaja.GrabaConfirmacionHabilitacionRemesa(CDate(txtFecha.Text), lsAgeCod, gsCodUser, gsOpeCod, Trim(txtGlosa.Text), Datos)
    Screen.MousePointer = 0
    
    If lbExito Then
        MsgBox "Se ha registrado satisfactoriamente la confirmación de la habilitaciones seleccionadas", vbInformation, "Aviso"
        'Set oImpre = New clases.NContImprimir
        'For i = 1 To UBound(Datos, 2)
        '    lsCadImpre = lsCadImpre & oImpre.ImprimeAsientoContable(Datos(5, i), gnLinPage, gnColPage, "CONFIRMACIÓN DE HABILITACIÓN REMESA", , "179") & oImpresora.gPrnSaltoPagina
        'Next
        'Set oImpre = Nothing
        'Set oPrevio = New clsprevio
        'oPrevio.Show lsCadImpre, "CONFIRMACIÓN DE HABILITACIÓN REMESA", False, gnLinPage
        'Set oPrevio = Nothing
        'cmdCancelar_Click
        chkTodos.value = 1 'Para reordenar columna
        EliminaItems Datos
        txtGlosa.Text = ""
        chkTodos.value = 0
    Else
        MsgBox "Ha sucedido un error al confirmar las Habilitaciones, si el error persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If
    cmdConfirmar.Enabled = True
    Set oCaja = Nothing
    Set oNContFunciones = Nothing
    Exit Sub
ErrCmdConfirmar:
    Screen.MousePointer = 0
    cmdConfirmar.Enabled = True
    MsgBox err.Description, vbCritical, "Aviso"
    
End Sub
Private Sub cmdRechazar_Click()
    Dim oCaja As COMNCajaGeneral.NCOMCajaGeneral
    Dim oNContFunciones As clases.NContFunciones
    Dim Datos As Variant
    Dim lsAgeCod As String
    Dim i As Integer
    Dim lbExito As Boolean
    Dim lsMovNro As String
    Dim lsCadImpre As String
    
    On Error GoTo ErrCmdConfirmar
    
    lsAgeCod = Right(txtAreaAgeCod.Text, 2)
    Datos = DameListaMovimientos(lsAgeCod)
    If UBound(Datos, 2) = 0 Then
        MsgBox "Ud. debe seleccionar al menos una Habilitación para continuar", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(Trim(txtGlosa.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la glosa del rechazo de la Habilitación", vbInformation, "Aviso"
        EnfocaControl txtGlosa
        Exit Sub
    End If
    
    If MsgBox("¿Esta seguro de rechazar la Habilitación de la remesa?", vbYesNo + vbQuestion, "Aviso") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    cmdRechazar.Enabled = False
    Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
    Set oNContFunciones = New clases.NContFunciones
    lsMovNro = oNContFunciones.GeneraMovNro(CDate(txtFecha.Text), lsAgeCod, gsCodUser)
    
    lbExito = oCaja.GrabaRechazoHabilitacionRemesa(lsMovNro, Trim(txtGlosa.Text), Datos)
    Screen.MousePointer = 0
    
    If lbExito Then
        MsgBox "Se ha rechazado satisfactoriamente las habilitaciones seleccionadas", vbInformation, "Aviso"
        'cmdCancelar_Click
        chkTodos.value = 1 'Para reordenar columna
        EliminaItems Datos
        txtGlosa.Text = ""
        chkTodos.value = 0
    Else
        MsgBox "Ha sucedido un error al rechazar las Habilitaciones, si el error persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If
    cmdRechazar.Enabled = True
    Set oCaja = Nothing
    Exit Sub
ErrCmdConfirmar:
    Screen.MousePointer = 0
    cmdRechazar.Enabled = True
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
Private Sub EliminaItems(ByRef pDatos As Variant)
    Dim i As Integer
    If IsArray(pDatos) Then
        For i = 1 To UBound(pDatos, 2)
            fg.EliminaFila CInt(pDatos(9, i)), True
        Next
    End If
End Sub
