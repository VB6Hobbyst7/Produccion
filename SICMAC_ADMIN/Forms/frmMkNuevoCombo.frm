VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMkNuevoCombo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nuevo Combo de Campaña"
   ClientHeight    =   6000
   ClientLeft      =   9750
   ClientTop       =   5085
   ClientWidth     =   8475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMkNuevoCombo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8475
   Begin VB.TextBox txtCantidad 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   295
      Left            =   5640
      TabIndex        =   21
      Top             =   2400
      Visible         =   0   'False
      Width           =   1510
   End
   Begin VB.CommandButton cmdExaProductos 
      Caption         =   "..."
      Height          =   300
      Left            =   6840
      TabIndex        =   20
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox textObjDes 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   295
      Left            =   5640
      TabIndex        =   19
      Top             =   2040
      Visible         =   0   'False
      Width           =   1510
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHProductos 
      Height          =   2655
      Left            =   240
      TabIndex        =   18
      Top             =   1200
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   7
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   360
      Left            =   7320
      TabIndex        =   17
      Top             =   5520
      Width           =   990
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   6240
      TabIndex        =   16
      Top             =   5520
      Width           =   990
   End
   Begin VB.TextBox txtRmaxDolares 
      Height          =   285
      Left            =   4920
      TabIndex        =   15
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox txtRmaxSoles 
      Height          =   285
      Left            =   4920
      TabIndex        =   14
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox txtRminDolares 
      Height          =   285
      Left            =   3840
      TabIndex        =   12
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox txtRminSoles 
      Height          =   285
      Left            =   3840
      TabIndex        =   10
      Top             =   4560
      Width           =   735
   End
   Begin VB.CheckBox chkDolares 
      Caption         =   "Dolares"
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CheckBox chkSoles 
      Caption         =   "Soles"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CheckBox chkAperturas 
      Caption         =   "Aperturas"
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CheckBox chkDesembolsos 
      Caption         =   "Desembolsos"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "Quitar"
      Height          =   360
      Left            =   1080
      TabIndex        =   3
      Top             =   720
      Width           =   990
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   195
      Left            =   4680
      TabIndex        =   13
      Top             =   5040
      Width           =   180
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   195
      Left            =   4680
      TabIndex        =   11
      Top             =   4560
      Width           =   180
   End
   Begin VB.Label lblMoneda 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Moneda:"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   4560
      Width           =   630
   End
   Begin VB.Label lblAplicableA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aplicable a Operaciones:"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   4080
      Width           =   1770
   End
   Begin VB.Label lblProductos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Productos"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   720
   End
   Begin VB.Label lblDescripcion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   810
   End
End
Attribute VB_Name = "frmMkNuevoCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oNGastosMarketing As New NGastosMarketing
Private Const Nuevo = 0
Private Const Editar = 1
Private Const Detalle = 3

Dim idCampana As String
Dim idcombo As String
Dim accion As Integer
Dim msjagregar As String
Private Type itemProducto
    codigo As String
    descripcion As String
    unidad As String
End Type
Public Sub aDetalle(pNIdCombo As String, nIdCampana As String, cComboDescripcion As String, bDesembolso As Integer, bApertura As Integer, _
                    bSoles As Integer, bDolares As Integer, nMinSoles As String, nMaxSoles As String, nMinDolares As String, nMaxDolares As String)
    
    
        'Deshabilitamos todo
    txtDescripcion.Enabled = False
    chkDesembolsos.Enabled = False
    chkAperturas.Enabled = False
    chkSoles.Enabled = False
    chkDolares.Enabled = False
    txtRminSoles.Enabled = False
    txtRmaxSoles.Enabled = False
    txtRminDolares.Enabled = False
    txtRmaxDolares.Enabled = False
    MSHProductos.Enabled = False
    cmdAceptar.Enabled = False
    cmdQuitar.Enabled = False
    
    aEditar pNIdCombo, nIdCampana, cComboDescripcion, bDesembolso, bApertura, bSoles, bDolares, nMinSoles, nMaxSoles, nMinDolares, nMaxDolares, "Detalle de Combo de Campaña" 'reutilizamos el editar
    accion = Detalle
    
End Sub

Public Sub aNuevo(pIdCampana As String)
    Me.Caption = "Nuevo Combo de Campaña"
    accion = Nuevo
    idCampana = pIdCampana
    Show 1
End Sub

Public Sub aEditar(pNIdCombo As String, nIdCampana As String, cComboDescripcion As String, bDesembolso As Integer, bApertura As Integer, _
                    bSoles As Integer, bDolares As Integer, nMinSoles As String, nMaxSoles As String, nMinDolares As String, nMaxDolares As String, ByVal sCaption As String)
    Me.Caption = sCaption
    accion = Editar
    idcombo = pNIdCombo
    idCampana = nIdCampana
    txtDescripcion.Text = cComboDescripcion
    chkDesembolsos.value = bDesembolso
    chkAperturas.value = bApertura
    chkSoles.value = bSoles
    chkDolares.value = bDolares
    txtRminSoles.Text = nMinSoles
    txtRmaxSoles.Text = nMaxSoles
    txtRminDolares.Text = nMinDolares
    txtRmaxDolares.Text = nMaxDolares
    'MSHProductos.Clear
    llenarComboBienes
    
    Show 1
End Sub

Private Sub cmdAgregar_Click()
    Dim lsRaiz As String
    Dim oDescObj As ClassDescObjeto

End Sub



Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdExaProductos_Click()
    Dim oConst As DConstantes
    Dim rs As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Set oConst = New DConstantes
    Dim oDesc As New ClassDescObjeto
    Set rs = New ADODB.Recordset
    
    Set rs = oNGastosMarketing.getMaterialesPromocion
    oDesc.lbUltNivel = True
    Set rsTemp = rs.Clone
    oDesc.Show rs, textObjDes
   
    If oDesc.lbOk Then
        AgregarProductoAGrilla rsTemp, MSHProductos.row, oDesc.gsSelecCod
    End If
End Sub

Private Sub cmdQuitar_Click()

    're vapi
    If (MSHProductos.TextMatrix(MSHProductos.RowSel, 1) = "" Or MSHProductos.TextMatrix(MSHProductos.RowSel, 1) = msjagregar) Then
        Exit Sub
    End If
    're vapi
    
    Dim pregunta As String
    pregunta = MsgBox("¿Esta Seguro que Quiere Eliminar el Producto?.", vbYesNo + vbExclamation + vbDefaultButton2, "Eliminar Producto.")
    If pregunta = vbYes Then
        GridEliminarFila MSHProductos, MSHProductos.RowSel
    End If
    
End Sub

Sub GridEliminarFila(Grid As Object, Fila As Single)
   
   
    Dim nIdComboBienes As String: nIdComboBienes = MSHProductos.TextMatrix(MSHProductos.RowSel, 6)
   
   '-- Si solo tiene una fila, la limpia en vez de borrarla
   If Grid.Rows = 2 Then
     Dim Columna As Single
     For Columna = 0 To Grid.Cols - 1
       Grid.TextMatrix(1, Columna) = ""
     Next Columna
   Else
     Grid.RemoveItem (Fila)
   End If
   
   
    If accion = Editar Then
        If nIdComboBienes <> "" Then
            oNGastosMarketing.EliminaBienEnCombo nIdComboBienes
        End If
        'If MSHProductos.RowSel = MSHProductos.Rows - 1 And nIdComboBienes = "" Then
        '    'MSHProductos.AddItem ""
        '    MSHProductos.Rows = MSHProductos.Rows + 1 're vapi
        '    MSHProductos.TextMatrix(MSHProductos.Rows - 1, 1) = msjagregar
       ' End If
    End If
   
End Sub

Private Sub Form_Load()
    msjagregar = "¡click aqui!"
    FormatoGrilla
    MSHProductos.TextMatrix(MSHProductos.Rows - 1, 1) = msjagregar
End Sub

Private Sub MSHProductos_Click()
    If MSHProductos.col = 1 Then
        EnfocaTexto textObjDes, 0, MSHProductos
        mostrarExaminadorProductos
    End If
    
    If (MSHProductos.col) = 4 And (MSHProductos.TextMatrix(MSHProductos.row, 1) <> msjagregar) Then
        mostrarCantidad
    End If
End Sub

Private Sub MSHProductos_GotFocus()
    ocultarExaminadorProductos
    ocultarCantidad
End Sub

Private Sub mostrarCantidad()
    MSHProductos.col = 4
    EnfocaTexto txtCantidad, 0, MSHProductos
End Sub

Private Sub mostrarExaminadorProductos()
    cmdExaProductos.Visible = True
    cmdExaProductos.Top = textObjDes.Top
    cmdExaProductos.Left = textObjDes.Left + textObjDes.Width - cmdExaProductos.Width
End Sub

Private Sub ocultarExaminadorProductos()
    textObjDes.Visible = False
    cmdExaProductos.Visible = False
End Sub
Private Sub FormatoGrilla()
    'captions
    MSHProductos.TextMatrix(0, 0) = ""
    MSHProductos.TextMatrix(0, 1) = "Objeto"
    MSHProductos.TextMatrix(0, 2) = "Descripción"
    MSHProductos.TextMatrix(0, 3) = "Unidad"
    MSHProductos.TextMatrix(0, 4) = "Cantidad"
    MSHProductos.TextMatrix(0, 5) = "nIdCombo"
    MSHProductos.TextMatrix(0, 6) = "nIdComboBienes"
    'tamaños
    MSHProductos.ColWidth(0) = 335
    MSHProductos.ColWidth(1) = 1500
    MSHProductos.ColWidth(2) = 2400
    MSHProductos.ColWidth(3) = 1500
    MSHProductos.ColWidth(4) = 1500
    MSHProductos.ColWidth(5) = 0
    MSHProductos.ColWidth(6) = 0
    
End Sub

Private Sub llenarComboBienes()
    Dim rs As ADODB.Recordset
    Set rs = oNGastosMarketing.RecuperaComboBienes(idcombo)
    Dim n As Integer
    n = 0
    're vapi
    Call LimpiaFlex(MSHProductos)
    FormatoGrilla
    'end re vapi
    Do While Not rs.EOF
        MSHProductos.AddItem ""
        MSHProductos.TextMatrix(n + 1, 6) = rs!nIdComboBienes
        MSHProductos.TextMatrix(n + 1, 5) = rs!nIdCombo
        MSHProductos.TextMatrix(n + 1, 1) = rs!cBSCod
        MSHProductos.TextMatrix(n + 1, 2) = rs!descripcion
        MSHProductos.TextMatrix(n + 1, 3) = rs!unidad
        MSHProductos.TextMatrix(n + 1, 4) = rs!nCantidad
        n = n + 1
        rs.MoveNext
    Loop
    MSHProductos.TextMatrix(MSHProductos.Rows - 1, 1) = msjagregar
End Sub

Private Sub AgregarProductoAGrilla(ByVal rs As ADODB.Recordset, ByVal pos As Integer, ByVal codigo)
    Dim item As itemProducto
    item = ObtenerProductoPorCodigo(codigo, rs)
    
    If Not BuscarRepetidos(MSHProductos, item.codigo) Then
        MSHProductos.TextMatrix(pos, 1) = Trim(item.codigo)
        MSHProductos.TextMatrix(pos, 2) = Trim(item.descripcion)
        MSHProductos.TextMatrix(pos, 3) = Trim(item.unidad)
        mostrarCantidad
        ocultarExaminadorProductos
    End If
End Sub

Private Function BuscarRepetidos(ByVal msflexgrid As Object, ByVal codigo As String) As Boolean
    Dim i As Integer
    For i = 1 To msflexgrid.Rows - 1
           If msflexgrid.TextMatrix(i, 1) = codigo Then
                BuscarRepetidos = True
                MsgBox "Se ha encontrado que ya tiene un registro con el mismo producto"
                Exit Function
           End If
    Next i
End Function

Private Function ObtenerProductoPorCodigo(ByVal codigo As String, ByVal rs As ADODB.Recordset) As itemProducto
    Dim item As itemProducto
    Do While Not rs.EOF
        If rs!codigo = codigo Then
            item.codigo = rs!codigo
            item.descripcion = rs!descripcion
            item.unidad = rs!unidad
            ObtenerProductoPorCodigo = item
            Exit Function
        End If
        rs.MoveNext
    Loop
End Function

Private Sub ocultarCantidad()
    txtCantidad.Visible = False
End Sub


Private Sub txtCantidad_LostFocus()
    If MSHProductos.TextMatrix(MSHProductos.Rows - 1, 4) = "" Then
        Dim i As Integer
        For i = 1 To MSHProductos.Cols - 1
            MSHProductos.TextMatrix(MSHProductos.Rows - 1, i) = ""
        Next i
        MSHProductos.TextMatrix(MSHProductos.Rows - 1, 1) = msjagregar
    End If
End Sub

Private Sub txtCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
're vapi
    Dim cBSCod As String
    Dim cantidad As String
    If KeyCode = 13 Then
        ocultarCantidad
        If Not IsNumeric(txtCantidad.Text) Or Val(txtCantidad.Text) <= 0 Then
            'txtCantidad.Text = 1
            MsgBox "Se dede ingresar valores numéricos mayores de cero", vbInformation, "¡Aviso!"
            If MSHProductos.Rows - 1 = MSHProductos.row Then
                Dim i As Integer
                For i = 1 To MSHProductos.Cols - 1
                    MSHProductos.TextMatrix(MSHProductos.Rows - 1, i) = ""
                Next i
                 MSHProductos.TextMatrix(MSHProductos.Rows - 1, 1) = msjagregar
            End If
            Exit Sub
        End If
        MSHProductos.TextMatrix(MSHProductos.row, 4) = txtCantidad.Text
        If (MSHProductos.row = MSHProductos.Rows - 1) And MSHProductos.col = 4 And (txtCantidad <> "") And (MSHProductos.TextMatrix(MSHProductos.row, 1) <> "") _
        Then 'es un nuevo registro
            If accion = Editar Then
                cBSCod = MSHProductos.TextMatrix(MSHProductos.row, 1)
                cantidad = MSHProductos.TextMatrix(MSHProductos.row, 4)
                registrarComboBienEditar idcombo, cBSCod, cantidad
            End If
           MSHProductos.AddItem ""
           MSHProductos.TextMatrix(MSHProductos.Rows - 1, 1) = msjagregar
        Else
            If accion = Editar Then 'es un editar un combo bien registro
                cBSCod = MSHProductos.TextMatrix(MSHProductos.row, 1)
                cantidad = MSHProductos.TextMatrix(MSHProductos.row, 4)
                If cBSCod = "" Then
                    MSHProductos.TextMatrix(MSHProductos.row, 4) = ""
                Else
                    ' aqui actualizamos el combo bien
                    Dim idCombobien As String: idCombobien = MSHProductos.TextMatrix(MSHProductos.row, 6)

                    oNGastosMarketing.ActualizaDetalleComboCon idCombobien, cBSCod, cantidad
                    
                End If
            End If
        End If
    End If
'end re vapi
End Sub






'validaciones del formulario

Private Sub registrarComboBienEditar(ByVal nIdCombo As String, ByVal cBSCod As String, ByVal nCantidad As String)
    Dim pregunta As String
    
    'Preguntamos si esta seguro de grabar
    pregunta = MsgBox("¿Está Seguro que va a registrar el bien?.", vbYesNo + vbExclamation + vbDefaultButton2, "Grabar Combo.")
    If pregunta <> vbYes Then
    're vapi
        MSHProductos.Rows = MSHProductos.Rows - 1
    'end revapi
        Exit Sub
    End If
    
    'grabamos el bien
    Dim nIdComboBienes As Integer
    nIdComboBienes = oNGastosMarketing.InsertaDetalleComboCon(nIdCombo, cBSCod, nCantidad)
    MSHProductos.TextMatrix(MSHProductos.row, 6) = nIdComboBienes 'este es el que se obtuvo al insertar
    MSHProductos.TextMatrix(MSHProductos.row, 5) = nIdCombo ' este es el parametro que llegó de la funcion
    
End Sub
'validaciones del formulario

Private Function estanMarcadosLosChecksMoneda() As Boolean
    If chkSoles.value = 1 Or chkDolares.value = 1 Then
        estanMarcadosLosChecksMoneda = True
        Exit Function
    End If
    estanMarcadosLosChecksMoneda = False
End Function


Private Function estanMarcadosLosChecksOperaciones() As Boolean
    If chkDesembolsos.value = 1 Or chkAperturas.value = 1 Then
        estanMarcadosLosChecksOperaciones = True
        Exit Function
    End If
    estanMarcadosLosChecksOperaciones = False
End Function

Private Function estaLlenoElGrid() As Boolean
    Dim dev As Boolean
    dev = True
    If MSHProductos.Rows = 2 Then
        If MSHProductos.TextMatrix(MSHProductos.Rows - 1, 4) = "" Then
            dev = False
        End If
    End If
    estaLlenoElGrid = dev
End Function

Private Function validarForm() As Boolean
    Dim mensaje As String
    Dim error As Boolean
    mensaje = ""
    error = False
    
    're vapi
    Dim i As Integer
    For i = 1 To MSHProductos.Rows - 1
        If MSHProductos.TextMatrix(i, 1) <> msjagregar Then
            Dim cBSCod As String: cBSCod = MSHProductos.TextMatrix(i, 1)
            Dim nCantidad As String: nCantidad = MSHProductos.TextMatrix(i, 4)
            If cBSCod = "" Or nCantidad = "" Then
                error = True
                mensaje = mensaje & vbNewLine & "- Ingrese la cantidad de todos los bienes merchandising"
            End If
        End If
    Next i
    
    'end revapi
    
    
    
    If ExisteDescripcionRepetida Then
        error = True
        mensaje = mensaje & vbNewLine & "- Parece que ya existe una descripción identica para otro combo"
    End If
    
    If txtDescripcion = "" Then
        error = True
        mensaje = mensaje & vbNewLine & "- Debe ingresar una descripcion"
    End If
    
    If Not estaLlenoElGrid Then
        error = True
        mensaje = mensaje & vbNewLine & "- Debe ingresar por lo menos un producto"
    End If
    
    If Not estanMarcadosLosChecksOperaciones Then
        error = True
        mensaje = mensaje & vbNewLine & "- Debe elegir a que operaciones se aplicaran los combos"
    End If
    
    If Not estanMarcadosLosChecksMoneda Then
        error = True
        mensaje = mensaje & vbNewLine & "- Debe elegir a que moneda se aplicaran los combos"
    End If
    
    If chkSoles.value = 1 Then
        If txtRminSoles.Text = "" Or txtRmaxSoles.Text = "" Then
            error = True
            mensaje = mensaje & vbNewLine & "- Debe llenar los intervalos para la moneda soles"
        Else
            If Val(txtRminSoles.Text) >= Val(txtRmaxSoles.Text) Then
                error = True
                mensaje = mensaje & vbNewLine & "- El valor mínimo del intervalo de soles debe ser menor que el valor máximo"
            End If
        End If
    End If
    
    If chkDolares.value = 1 Then
        If txtRminDolares.Text = "" Or txtRmaxDolares.Text = "" Then
            error = True
            mensaje = mensaje & vbNewLine & "- Debe llenar los intervalos para la moneda dólares"
        Else
            If Val(txtRminDolares.Text) >= Val(txtRmaxDolares.Text) Then
                error = True
                mensaje = mensaje & vbNewLine & "- El valor mínimo del intervalo de dólares debe ser menor que el valor máximo"
            End If
        End If
    End If
    
    
    
    If error Then
        MsgBox mensaje, vbOKOnly, "verifique los datos!"
    End If
    
    validarForm = Not error
End Function



Private Sub txtRmaxDolares_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 97) And (KeyAscii < 122) Or (KeyAscii >= 65) And (KeyAscii < 90) Then
        KeyAscii = 8
    End If
End Sub

Private Sub txtRmaxSoles_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 97) And (KeyAscii < 122) Or (KeyAscii >= 65) And (KeyAscii < 90) Then
        KeyAscii = 8
    End If
End Sub


Private Sub txtRminDolares_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 97) And (KeyAscii < 122) Or (KeyAscii >= 65) And (KeyAscii < 90) Then
        KeyAscii = 8
    End If
End Sub

Private Sub txtRminSoles_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 97) And (KeyAscii < 122) Or (KeyAscii >= 65) And (KeyAscii < 90) Then
      KeyAscii = 8
    End If
End Sub

Private Sub cmdAceptar_Click()
    If validarForm Then
        corregirIntervalosCheckMoneda
        Select Case accion
        Case Nuevo
            RegistrarCombo
            Unload Me
        Case Editar
            EditarCombo
        End Select
        
    End If
End Sub
Private Function ExisteDescripcionRepetida() As Boolean
    Dim cComboDescripcion As String: cComboDescripcion = txtDescripcion.Text
    Dim existe As Integer
    existe = oNGastosMarketing.ExisteDescripcionCombo(idcombo, idCampana, cComboDescripcion)
    If existe = 0 Then
        ExisteDescripcionRepetida = False
    Else
        ExisteDescripcionRepetida = True
    End If
End Function
Private Sub corregirIntervalosCheckMoneda()
    If chkSoles.value = 0 Then
        txtRminSoles.Text = ""
        txtRmaxSoles.Text = ""
    End If
    If chkDolares.value = 0 Then
        txtRminDolares.Text = ""
        txtRmaxDolares.Text = ""
    End If
End Sub
'Acciones a realizar
Private Sub EditarCombo()
    Dim pregunta As String
     'Preguntamos si esta seguro de grabar
    pregunta = MsgBox("¿Está Seguro que va a editar el combo?.", vbYesNo + vbExclamation + vbDefaultButton2, "Grabar Combo.")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    'Actualizamos los valores
    Dim cComboDescripcion As String: cComboDescripcion = txtDescripcion.Text
    Dim bDesembolso As Integer: bDesembolso = chkDesembolsos.value
    Dim bApertura As Integer: bApertura = chkAperturas.value
    Dim bSoles As Integer: bSoles = chkSoles.value
    Dim bDolares As Integer: bDolares = chkDolares.value
    Dim nMinSoles As String: nMinSoles = IIf(txtRminSoles.Text = "", "NULL", txtRminSoles.Text)
    Dim nMaxSoles As String: nMaxSoles = IIf(txtRmaxSoles.Text = "", "NULL", txtRmaxSoles.Text)
    Dim nMinDolares As String: nMinDolares = IIf(txtRminDolares.Text = "", "NULL", txtRminDolares.Text)
    Dim nMaxDolares As String: nMaxDolares = IIf(txtRmaxDolares.Text = "", "NULL", txtRmaxDolares.Text)
    
    oNGastosMarketing.ActualizaComboxCampana idcombo, idCampana, cComboDescripcion, bDesembolso, bApertura, bSoles, bDolares, nMinSoles, nMaxSoles, nMinDolares, nMaxDolares
    're vapi
    MsgBox "Se han actualizado los datos del combo por campaña!"
    'end re vapi
End Sub

Private Function RegistrarCombo()
    Dim pregunta As String
    Dim nIdCombo As Integer
    Dim oCon As DConecta
    'Preguntamos si esta seguro de grabar
    pregunta = MsgBox("¿Está Seguro que va a registrar el combo?.", vbYesNo + vbExclamation + vbDefaultButton2, "Grabar Combo.")
    If pregunta <> vbYes Then
        Exit Function
    End If
    'obtenemos los valores del formulario
    Dim cComboDescripcion As String: cComboDescripcion = txtDescripcion.Text
    Dim bDesembolso As Integer: bDesembolso = chkDesembolsos.value
    Dim bApertura As Integer: bApertura = chkAperturas.value
    Dim bSoles As Integer: bSoles = chkSoles.value
    Dim bDolares As Integer: bDolares = chkDolares.value
    Dim nMinSoles As String: nMinSoles = IIf(txtRminSoles.Text = "", "NULL", txtRminSoles.Text)
    Dim nMaxSoles As String: nMaxSoles = IIf(txtRmaxSoles.Text = "", "NULL", txtRmaxSoles.Text)
    Dim nMinDolares As String: nMinDolares = IIf(txtRminDolares.Text = "", "NULL", txtRminDolares.Text)
    Dim nMaxDolares As String: nMaxDolares = IIf(txtRmaxDolares.Text = "", "NULL", txtRmaxDolares.Text)
    'hacemos la operacion de registro de combo y sus detalles
    Set oCon = oNGastosMarketing.getOcon
    oCon.AbreConexion
    oCon.BeginTrans
    nIdCombo = oNGastosMarketing.InsertaComboCampana(idCampana, cComboDescripcion, bDesembolso, bApertura, bSoles, bDolares, nMinSoles, nMaxSoles, nMinDolares, nMaxDolares)
    If nIdCombo <> -1 Then
        Dim i As Integer
            For i = 1 To MSHProductos.Rows - 1
                If MSHProductos.TextMatrix(i, 1) <> msjagregar Then
                Dim cBSCod As String: cBSCod = MSHProductos.TextMatrix(i, 1)
                Dim nCantidad As String: nCantidad = MSHProductos.TextMatrix(i, 4)
                   Call oNGastosMarketing.InsertaDetalleCombo(nIdCombo, cBSCod, nCantidad)
                End If
            Next i
    End If
    oCon.CommitTrans
    oCon.CierraConexion
End Function





