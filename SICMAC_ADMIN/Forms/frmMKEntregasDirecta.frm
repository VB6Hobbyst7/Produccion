VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMKEntregasDirecta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entregas Directas"
   ClientHeight    =   9810
   ClientLeft      =   9765
   ClientTop       =   4515
   ClientWidth     =   9225
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMKEntregasDirecta.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9810
   ScaleWidth      =   9225
   Begin Sicmact.Usuario usuario 
      Left            =   1680
      Top             =   9000
      _extentx        =   820
      _extenty        =   820
   End
   Begin MSComCtl2.DTPicker txtFecha 
      Height          =   315
      Left            =   7560
      TabIndex        =   17
      Top             =   120
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      _Version        =   393216
      Format          =   63569921
      CurrentDate     =   37156
   End
   Begin Sicmact.Usuario empleado 
      Left            =   600
      Top             =   9000
      _extentx        =   820
      _extenty        =   820
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   360
      Left            =   8040
      TabIndex        =   16
      Top             =   9000
      Width           =   990
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   360
      Left            =   6960
      TabIndex        =   15
      Top             =   9000
      Width           =   990
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   855
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   7680
      Width           =   8775
   End
   Begin VB.Frame FraCliente 
      Caption         =   "Cliente"
      ClipControls    =   0   'False
      Height          =   975
      Left            =   240
      TabIndex        =   9
      Top             =   6480
      Width           =   8775
      Begin Sicmact.TxtBuscar TxtBuscarCli 
         Height          =   345
         Left            =   840
         TabIndex        =   10
         Top             =   360
         Width           =   1980
         _extentx        =   3493
         _extenty        =   609
         appearance      =   1
         appearance      =   1
         font            =   "frmMKEntregasDirecta.frx":030A
         appearance      =   1
         tipobusqueda    =   3
         stitulo         =   ""
         tipobuspers     =   1
      End
      Begin VB.Label lblNombreCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2880
         TabIndex        =   12
         Top             =   360
         Width           =   4755
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.ComboBox cmbCampana 
      Height          =   315
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3000
      Width           =   3135
   End
   Begin VB.ComboBox cmbTipoBusqueda 
      Height          =   315
      ItemData        =   "frmMKEntregasDirecta.frx":032E
      Left            =   1200
      List            =   "frmMKEntregasDirecta.frx":0338
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Solicitador por:"
      ClipControls    =   0   'False
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   8775
      Begin Sicmact.TxtBuscar txtBuscaPers 
         Height          =   345
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   1980
         _extentx        =   3493
         _extenty        =   609
         appearance      =   1
         appearance      =   1
         font            =   "frmMKEntregasDirecta.frx":034D
         appearance      =   1
         tipobusqueda    =   7
         stitulo         =   ""
         tipobuspers     =   1
         enabledtext     =   0
      End
      Begin VB.Label lblPersNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2880
         TabIndex        =   3
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label lblusuario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   600
      End
   End
   Begin Sicmact.FlexEdit flxProductos 
      Height          =   2415
      Left            =   240
      TabIndex        =   8
      Top             =   3840
      Width           =   8775
      _extentx        =   15478
      _extenty        =   4260
      cols0           =   5
      highlight       =   1
      allowuserresizing=   3
      rowsizingmode   =   1
      encabezadosnombres=   "#-Combo/Producto-idProducto-Cant.-Elegir"
      encabezadosanchos=   "500-6500-0-500-500"
      font            =   "frmMKEntregasDirecta.frx":0371
      font            =   "frmMKEntregasDirecta.frx":039D
      font            =   "frmMKEntregasDirecta.frx":03C9
      font            =   "frmMKEntregasDirecta.frx":03F5
      font            =   "frmMKEntregasDirecta.frx":0421
      fontfixed       =   "frmMKEntregasDirecta.frx":044D
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1
      columnasaeditar =   "X-X-X-3-4"
      textstylefixed  =   3
      listacontroles  =   "0-0-0-0-4"
      encabezadosalineacion=   "C-L-C-R-L"
      formatosedit    =   "0-0-0-3-0"
      textarray0      =   "#"
      lbeditarflex    =   -1
      lbbuscaduplicadotext=   -1
      colwidth0       =   495
      rowheight0      =   300
      forecolorfixed  =   -2147483630
   End
   Begin VB.Frame FraUsuaurio 
      Caption         =   "Entregado por: "
      Height          =   1215
      Left            =   240
      TabIndex        =   19
      Top             =   360
      Width           =   8775
      Begin VB.Label lblAgencia 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   840
         TabIndex        =   26
         Top             =   720
         Width           =   1980
      End
      Begin VB.Label lblAgenci 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agencia:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   630
      End
      Begin VB.Label lblCodUserReal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   840
         TabIndex        =   23
         Top             =   360
         Width           =   1980
      End
      Begin VB.Label lblNombreuserreal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2880
         TabIndex        =   22
         Top             =   360
         Width           =   4755
      End
      Begin VB.Label lblUsuarioreallb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   600
      End
      Begin VB.Label lblUsuarioReal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7680
         TabIndex        =   20
         Top             =   360
         Width           =   915
      End
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "fecha:"
      Height          =   195
      Left            =   6960
      TabIndex        =   24
      Top             =   120
      Width           =   465
   End
   Begin VB.Label lblCargandoProductos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando Productos..."
      Height          =   195
      Left            =   1200
      TabIndex        =   18
      Top             =   3360
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label lblDescripcion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Top             =   7440
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Campaña: "
      Height          =   195
      Left            =   3840
      TabIndex        =   6
      Top             =   3000
      Width           =   780
   End
   Begin VB.Label lblBuscarPor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar por:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   825
   End
End
Attribute VB_Name = "frmMKEntregasDirecta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oNGastosMarketing As New NGastosMarketing
Dim rsProductos As ADODB.Recordset
Private Type itemComboBox
    cod As String
    dsc As String
End Type
Dim lsCampanas() As itemComboBox
Private Sub cmbCampana_Click()
    Dim idCampana As String
    idCampana = getIdLista(cmbCampana.ListIndex, lsCampanas)
    llenarFlexCombos idCampana
End Sub


Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub flxProductos_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    If pnCol = 4 And pnRow <> 0 And flxProductos.TextMatrix(pnRow, 3) = "" Then
        flxProductos.TextMatrix(pnRow, 3) = 1
    End If
End Sub

Private Sub Form_Load()
    txtFecha = Format(gdFecSis, "dd/mm/yyyy")
    cmbCampana.Enabled = False
    llenarComboCampanasActivas
    lblUsuarioReal = gsCodUser
    
    usuario.DatosPers gsCodPersUser
    lblNombreuserreal = PstaNombre(usuario.UserNom)
    lblCodUserReal = usuario.PersCod
    lblAgencia = usuario.DescAgeAct

End Sub
Private Sub cmbTipoBusqueda_Click()
    If cmbTipoBusqueda = "Combo" Then
        cmbCampana.Enabled = True
        resetFlex True
        cmbCampana_Click
    End If
    If cmbTipoBusqueda = "Producto" Then
        resetFlex False
        cmbCampana = cmbCampana.List(0)
        cmbCampana.Enabled = False
        If rsProductos Is Nothing Then
            lblCargandoProductos.Visible = True
            Set rsProductos = oNGastosMarketing.getMaterialesPromocionConSaldoXalmacen(Val(gsCodAge))
            lblCargandoProductos.Visible = False
        End If
        llenarFlexProductos rsProductos
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsProductos = Nothing
End Sub



Private Sub txtBuscaPers_EmiteDatos()
    Call CargaDatosEmpleado(txtBuscaPers)
End Sub
're vapi
Public Function CargaDatosEmpleado(ByVal psPersCod As String) As Boolean
    CargaDatosEmpleado = False
    empleado.DatosPers psPersCod
    If psPersCod <> "" Then
        txtBuscaPers.Text = empleado.PersCod
        lblPersNombre = PstaNombre(empleado.UserNom)
    End If
End Function
're vapi
Private Sub TxtBuscarCli_EmiteDatos()
    Call CargaDatosCliente(TxtBuscarCli)
End Sub
Public Function CargaDatosCliente(ByVal psPersCod As String) As Boolean
    CargaDatosCliente = False
    Dim oPersona As DPersona
    Set oPersona = New DPersona
    Call oPersona.RecuperaPersona(Trim(psPersCod))
    If oPersona.PersCodigo = "" Then
        MsgBox "No se pudo encontrar los datos de la Persona," & Chr(10) & " Verifique que la Persona exista", vbInformation, "Aviso"
        Exit Function
    End If
    TxtBuscarCli.Text = psPersCod
    lblNombreCliente = PstaNombre(oPersona.NombreCompleto)
End Function
Private Sub llenarComboCampanasActivas()
    Dim rs As ADODB.Recordset
    Set rs = oNGastosMarketing.RecuperaCampanas
    Dim n As Integer
    n = 0
    Do While Not rs.EOF
        ReDim Preserve lsCampanas(n)
        lsCampanas(n).cod = rs!nConsValor
        lsCampanas(n).dsc = rs!cConsDescripcion
        cmbCampana.AddItem Trim(lsCampanas(n).dsc)
        n = n + 1
        rs.MoveNext
    Loop
    cmbCampana = cmbCampana.List(0)
End Sub
Private Function getIdLista(ByVal index As Integer, ByRef item() As itemComboBox) As String
    getIdLista = item(index).cod
End Function
Private Sub resetFlex(ByVal combo As Boolean)
    flxProductos.Clear
    flxProductos.Rows = 2
    flxProductos.TextMatrix(0, 0) = "#"
    flxProductos.TextMatrix(0, 3) = "Cant."
    flxProductos.TextMatrix(0, 4) = "Elegir"
    If combo Then
        flxProductos.TextMatrix(0, 1) = "Combo"
        flxProductos.ColWidth(3) = 0
    Else
        flxProductos.TextMatrix(0, 1) = "Producto"
        flxProductos.ColWidth(3) = 500
    End If
    
End Sub


Private Sub llenarFlexProductos(ByVal rsc As ADODB.Recordset)
    Dim rs As ADODB.Recordset
    Set rs = rsc.Clone
    Dim lnFila As Integer
    resetFlex False
    Do While Not rs.EOF
        flxProductos.AdicionaFila
        lnFila = flxProductos.row
        flxProductos.TextMatrix(lnFila, 1) = rs!Val
        flxProductos.TextMatrix(lnFila, 2) = rs!codigo
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub
Private Sub llenarFlexCombos(ByVal idCampana As String)
    Dim lnFila As Integer
    Dim rs As ADODB.Recordset
    Set rs = oNGastosMarketing.RecuperaCombosXCampana(idCampana)
    Dim n As Integer
    resetFlex True
    Do While Not rs.EOF
        flxProductos.AdicionaFila
        lnFila = flxProductos.row
        flxProductos.TextMatrix(lnFila, 1) = rs!cComboDescripcion
        flxProductos.TextMatrix(lnFila, 2) = rs!nIdCombo
        rs.MoveNext
    Loop
End Sub
'validaciones

Private Function HaSeleccionadoProducto()
    Dim i As Integer
    For i = 1 To flxProductos.Rows - 1
        If flxProductos.TextMatrix(i, 4) = "." Then
            HaSeleccionadoProducto = True
            Exit Function
        End If
    Next i
    HaSeleccionadoProducto = False
End Function

Private Function ValidarFormulario() As Boolean
    Dim mensaje As String
    Dim error As Boolean
    mensaje = ""
    error = False
    
'    If txtBuscaPers = "" Then
'        error = True
'        mensaje = mensaje & vbNewLine & "- Debe elegir un usuario"
'    End If
    
    If TxtBuscarCli = "" Then
        error = True
        mensaje = mensaje & vbNewLine & "- Debe elegir un cliente"
    End If
    
    If txtDescripcion = "" Then
        error = True
        mensaje = mensaje & vbNewLine & "- Parece que no ha especificado una descripción"
    End If
    
    If Not HaSeleccionadoProducto Then
        error = True
        mensaje = mensaje & vbNewLine & "- Parece que no ha seleccionado ningún producto o combo"
    End If
    
    If error Then
        MsgBox mensaje, vbOKOnly, "¡Oops, verifique los datos!"
    End If
    
    ValidarFormulario = Not error
    
End Function

'accion de guardar
Private Sub cmdAceptar_Click()
    Dim pregunta As String
    'Preguntamos si esta seguro de grabar
    pregunta = MsgBox("¿Está Seguro que va a registrar la entrega?.", vbYesNo + vbExclamation + vbDefaultButton2, "Grabar Entrega.")
    If pregunta <> vbYes Then
        Exit Sub
    End If

    If ValidarFormulario Then
        If cmbTipoBusqueda = "Combo" Then
            registrarEntregaCombos
        Else
            registrarEntregaProductos
        End If
        MsgBox "¡Se ha registrado la entrega!"
        limpiarForm
    End If

End Sub

Private Sub limpiarForm()
    txtBuscaPers = ""
    lblPersNombre = ""
    resetFlex False
    TxtBuscarCli.Text = ""
    lblNombreCliente = ""
    txtDescripcion.Text = ""
End Sub

Private Sub registrarEntregaCombos()
    
    Dim oMov As DMov
    Set oMov = New DMov
    Dim sMovNro As String: sMovNro = oMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    
    
    Dim oCon As DConecta
    Dim i As Integer
    Set oCon = oNGastosMarketing.getOcon
            
        
    oCon.AbreConexion
    oCon.BeginTrans
    Dim idEntrega As Integer


    'los datos a ingresar de la entrega
        Dim cAgencia As String: cAgencia = usuario.CodAgeAct
        
        Dim cPersCodUser As String: cPersCodUser = lblCodUserReal
        Dim cPersCodUserSol As String: cPersCodUserSol = empleado.PersCod
        
        Dim cPersCodCliente As String: cPersCodCliente = TxtBuscarCli.Text
        Dim fecha As String: fecha = Format(txtFecha.value, "yyyymmdd")
        Dim cGlosa As String: cGlosa = txtDescripcion.Text
        idEntrega = oNGastosMarketing.InsertaEntregaCampana(cAgencia, cPersCodUser, cPersCodUserSol, cPersCodCliente, fecha, cGlosa, sMovNro)
        Dim nIdCampana As Integer: nIdCampana = Val(getIdLista(cmbCampana.ListIndex, lsCampanas))


    For i = 1 To flxProductos.Rows - 1

        If flxProductos.TextMatrix(i, 4) = "." Then
        
            Dim idcombo As String: idcombo = flxProductos.TextMatrix(i, 2)
            Dim rs As ADODB.Recordset
            Set rs = oNGastosMarketing.RecuperaComboBienesInserta(idcombo)
            
            Do While Not rs.EOF
                Call oNGastosMarketing.InsertaDetalleEntregaCampana(idEntrega, rs!cBSCod, idcombo, nIdCampana, rs!nCantidad, 1)
                rs.MoveNext
            Loop
            
        End If

    Next i
    oCon.CommitTrans
    oCon.CierraConexion
End Sub

Private Sub registrarEntregaProductos()
    Dim i As Integer
    Dim oMov As DMov
    Set oMov = New DMov
    Dim sMovNro As String: sMovNro = oMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    
    Dim oCon As DConecta
    Set oCon = oNGastosMarketing.getOcon
    oCon.AbreConexion
    oCon.BeginTrans
    
    'los datos a ingresar de la entrega
    Dim cAgencia As String: cAgencia = usuario.CodAgeAct
    
    Dim cPersCodUser As String: cPersCodUser = lblCodUserReal
    Dim cPersCodUserSol As String: cPersCodUserSol = empleado.PersCod
    
    
    Dim cPersCodCliente As String: cPersCodCliente = TxtBuscarCli.Text
    Dim fecha As String: fecha = Format(txtFecha.value, "yyyymmdd")
    Dim cGlosa As String: cGlosa = txtDescripcion.Text
    Dim idEntrega As Integer
    idEntrega = oNGastosMarketing.InsertaEntregaCampana(cAgencia, cPersCodUser, cPersCodUserSol, cPersCodCliente, fecha, cGlosa, sMovNro)
    For i = 1 To flxProductos.Rows - 1
        If flxProductos.TextMatrix(i, 4) = "." Then
            Dim cBSCod As String: cBSCod = flxProductos.TextMatrix(i, 2)
            Dim nCantidad As Integer: nCantidad = Val(flxProductos.TextMatrix(i, 3))
            Call oNGastosMarketing.InsertaDetalleEntregaCampana(idEntrega, cBSCod, "NULL", "NULL", nCantidad, 0)
        End If

    Next i
    oCon.CommitTrans
    oCon.CierraConexion
End Sub


