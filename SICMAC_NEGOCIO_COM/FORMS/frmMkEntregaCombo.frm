VERSION 5.00
Begin VB.Form frmMkEntregaCombo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrega de Merchandising"
   ClientHeight    =   3135
   ClientLeft      =   11310
   ClientTop       =   5955
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMkEntregaCombo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   6750
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   360
      Left            =   2760
      TabIndex        =   1
      Top             =   2400
      Width           =   1095
   End
   Begin SICMACT.FlexEdit flxCombos 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   3836
      Cols0           =   4
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-idCombo-Combo de Productos-Elegir"
      EncabezadosAnchos=   "500-0-4800-500"
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
      ColumnasAEditar =   "X-X-X-3"
      ListaControles  =   "0-0-0-4"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-C"
      FormatosEdit    =   "0-0-0-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   495
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmMkEntregaCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oNGastosMarketing As New NGastosMarketing
Dim idCampana As String
Dim sOperacion As String
Dim cPersCod As String
Dim rsCombos As ADODB.Recordset
Dim bDesembolso As Integer
'Dim rsProductos As ADODB.Recordset

'Private Sub flxProductos_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
'    If pnCol = 4 And pnRow <> 0 And flxProductos.TextMatrix(pnRow, 3) = "" Then
'        flxProductos.TextMatrix(pnRow, 3) = 1
'    End If
'End Sub


Private Sub Form_Load()
'    SSTTipo.Tab = 0
    llenarCombosPorCampana
End Sub
Public Sub inicio(ByVal cCtaCod As String, ByVal PsOperacion As String, ByVal esDesembolso As Boolean, ByVal Moneda As Integer, ByVal Monto)
    Dim rs As ADODB.Recordset
        Set rs = oNGastosMarketing.RecuperaCampanaPorCuenta(cCtaCod, esDesembolso)
        If rs.EOF Then
            Exit Sub
        End If
        
        bDesembolso = IIf(esDesembolso, 1, 0)
        idCampana = rs!idCampana
        cPersCod = rs!cPersCod
        sOperacion = PsOperacion
        If idCampana = "-1" Then
            Exit Sub
        End If
        Set rsCombos = oNGastosMarketing.RecuperaCombosXCampanaCondicion(idCampana, esDesembolso, Moneda, Monto)
        If rsCombos.EOF Then
            Exit Sub
        End If
        Me.Show 1
End Sub
Private Function llenarCombosPorCampana()
    Dim n As Integer
    n = 0
    flxCombos.Rows = 2
    flxCombos.Clear
    flxCombos.TextMatrix(1, 2) = ""
    Do While Not rsCombos.EOF
        flxCombos.AdicionaFila
        flxCombos.TextMatrix(n + 1, 1) = rsCombos!nIdCombo
        flxCombos.TextMatrix(n + 1, 2) = rsCombos!cComboDescripcion
        n = n + 1
        rsCombos.MoveNext
    Loop

End Function

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
    Dim cGlosa As String
    
Select Case sOperacion

            Case "100101"
                 cGlosa = "Créditos Desembolso en Efectivo"
            Case "100102"
                 cGlosa = "Créditos Desembolso Abono a Cuenta"
            Case "200101"
                 cGlosa = "Ahorros Apertura Efectivo"
            Case "200103"
                 cGlosa = "Ahorros Apertura Transferencia Banco"
            Case "210101"
                 cGlosa = "Plazo Fijo Apertura Efectivo"
            Case "210102"
                 cGlosa = "Plazo Fijo Apertura Cheque"
            Case "210103"
                 cGlosa = "Plazo Fijo Apertura Transferencia Banco"

    End Select
    
    idEntrega = oNGastosMarketing.InsertaEntregaCampana(Right(gsCodAge, 2), gsCodPersUser, cPersCod, Format(gdFecSis, "yyyymmdd"), cGlosa, sMovNro, bDesembolso)

    For i = 1 To flxCombos.Rows - 1

        If flxCombos.TextMatrix(i, 3) = "." Then
        
            Dim idcombo As String: idcombo = flxCombos.TextMatrix(i, 1)
            Dim rs As ADODB.Recordset
            Set rs = oNGastosMarketing.RecuperaComboBienesInserta(idcombo)
            
            Do While Not rs.EOF
                Call oNGastosMarketing.InsertaDetalleEntregaCampana(idEntrega, rs!cBSCod, idcombo, idCampana, rs!nCantidad, 1)
                rs.MoveNext
            Loop
            
        End If

    Next i
    oCon.CommitTrans
    oCon.CierraConexion
    MsgBox "¡Se ha Realizado de entregas de combos Merchandising, Tenga un buen dia!"
End Sub

Private Function HaSeleccionadoCombo()
    Dim i As Integer
    For i = 1 To flxCombos.Rows - 1
        If flxCombos.TextMatrix(i, 3) = "." Then
            HaSeleccionadoCombo = True
            Exit Function
        End If
    Next i
    HaSeleccionadoCombo = False
End Function
'Private Function HaSeleccionadoProducto()
'    Dim i As Integer
'    For i = 1 To flxProductos.Rows - 1
'        If flxProductos.TextMatrix(i, 4) = "." Then
'            HaSeleccionadoProducto = True
'            Exit Function
'        End If
'    Next i
'    HaSeleccionadoProducto = False
'End Function
Private Sub cmdAceptar_Click()
'    If SSTTipo.Tab = 0 Then
        If HaSeleccionadoCombo Then
            registrarEntregaCombos
            Unload Me
        Else
            Unload Me
        End If
'    End If
'    If SSTTipo.Tab = 1 Then
'        If HaSeleccionadoProducto Then
'            registrarEntregaProductos
'            Unload Me
'        Else
'            Unload Me
'        End If
'    End If
    
End Sub

'Private Sub registrarEntregaProductos()
'     Dim oMov As DMov
'    Set oMov = New DMov
'    Dim sMovNro As String: sMovNro = oMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
'    Dim oCon As DConecta
'    Dim i As Integer
'    Set oCon = oNGastosMarketing.getOcon
'    oCon.AbreConexion
'    oCon.BeginTrans
'    Dim idEntrega As Integer
'    'los datos a ingresar de la entrega
'    Dim cGlosa As String
'
'    Select Case sOperacion
'
'            Case "100101"
'                 cGlosa = "Créditos Desembolso en Efectivo"
'            Case "100102"
'                 cGlosa = "Créditos Desembolso Abono a Cuenta"
'            Case "200101"
'                 cGlosa = "Ahorros Apertura Efectivo"
'            Case "200103"
'                 cGlosa = "Ahorros Apertura Transferencia Banco"
'            Case "210101"
'                 cGlosa = "Plazo Fijo Apertura Efectivo"
'            Case "210102"
'                 cGlosa = "Plazo Fijo Apertura Cheque"
'            Case "210103"
'                 cGlosa = "Plazo Fijo Apertura Transferencia Banco"
'
'    End Select
'
'    idEntrega = oNGastosMarketing.InsertaEntregaCampana(Right(gsCodAge, 2), gsCodPersUser, cPersCod, Format(gdFecSis, "yyyymmdd"), cGlosa, sMovNro, bDesembolso)
'
'    For i = 1 To flxProductos.Rows - 1
'
'        If flxProductos.TextMatrix(i, 4) = "." Then
'
'            Dim idProducto As String: idProducto = flxProductos.TextMatrix(i, 1)
'            Dim nCantidad As Integer: nCantidad = flxProductos.TextMatrix(i, 3)
'
'            Call oNGastosMarketing.InsertaDetalleEntregaCampana(idEntrega, idProducto, "NULL", "NULL", nCantidad, 0)
'
'
'        End If
'
'    Next i
'    oCon.CommitTrans
'    oCon.CierraConexion
'    MsgBox "¡Se ha Realizado de entregas de Productos Merchandising, Tenga un buen dia!"
'End Sub


'Private Sub Form_Unload(Cancel As Integer)
'    Set rsProductos = Nothing
'End Sub

'Private Sub SSTTipo_Click(PreviousTab As Integer)
'    If SSTTipo.Tab = 1 And rsProductos Is Nothing Then
'        Dim pregunta As String
'        'Preguntamos si esta seguro de cargar los productos  del almacen
'        pregunta = MsgBox("¿Está Seguro que va a entregar productos sin combo?", vbYesNo + vbExclamation + vbDefaultButton2, "Confirmar.")
'        If pregunta <> vbYes Then
'            SSTTipo.Tab = 0
'            Exit Sub
'        End If
'        llenarFlexProductos
'    End If
'End Sub
'Private Sub llenarFlexProductos()
'
'    Set rsProductos = oNGastosMarketing.getMaterialesPromocionConSaldoXalmacen(val(gsCodAge))
'    Dim lnFila As Integer
'    Do While Not rsProductos.EOF
'        flxProductos.AdicionaFila
'        lnFila = flxProductos.row
'        flxProductos.TextMatrix(lnFila, 2) = rsProductos!val
'        flxProductos.TextMatrix(lnFila, 1) = rsProductos!codigo
'        rsProductos.MoveNext
'    Loop
'    rsProductos.Close
'    'Set rsProductos = Nothing
'End Sub
