VERSION 5.00
Begin VB.Form frmRubrosPagosAdelantados 
   Caption         =   "Rubros"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6510
   Icon            =   "frmRubrosPagosAdelantados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelarRubro 
      Caption         =   "Ca&ncelar"
      Height          =   350
      Left            =   5280
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptarRubro 
      Caption         =   "A&ceptar"
      Height          =   350
      Left            =   5280
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdEditarRubro 
      Caption         =   "E&ditar"
      Height          =   350
      Left            =   5280
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdNuevoRubro 
      Caption         =   "&Nuevo"
      Height          =   350
      Left            =   5280
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdEliminarRubro 
      Caption         =   "E&liminar"
      Default         =   -1  'True
      Height          =   350
      Left            =   5280
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin Sicmact.FlexEdit FERubro 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3201
      Cols0           =   5
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Nro-Rubro-Fórm. Cta. Cont. Debe.-nConsValor-IdRubroPagoAdelantadoCab"
      EncabezadosAnchos=   "400-2500-1800-3-3"
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
      ColumnasAEditar =   "X-1-2-X-X"
      ListaControles  =   "0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-R-R"
      FormatosEdit    =   "0-1-0-3-3"
      TextArray0      =   "Nro"
      lbFlexDuplicados=   0   'False
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmRubrosPagosAdelantados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'***Nombre:         frmPagosAdelantados
'***Descripción:    Formulario que permite el registro del
'                   pago adelantado.
'***Creación:       ELRO el 20111109 según Acta 323-2011/TI-D
'************************************************************
Option Explicit

Private Enum Accion
gValorDefectoAccion = 0
gNuevoRegistro = 1
gEditarRegistro = 2
gEliminarRegistro = 3
End Enum

Private fnAccion, fnFilaNoEditar As Integer

Private Sub CargarFERubro()
    Dim oDDocumento As DDocumento
    Set oDDocumento = New DDocumento
    Dim rsRubros As ADODB.Recordset
    Set rsRubros = New ADODB.Recordset
    Dim i As Integer
    fnAccion = gValorDefectoAccion
    fnFilaNoEditar = -1

    Set rsRubros = oDDocumento.recuperarRubroPagosAdelantados()
    
    Call LimpiaFlex(FERubro)
    
    FERubro.lbEditarFlex = True
    
    If Not rsRubros.BOF And Not rsRubros.EOF Then
    i = 1
        Do While Not rsRubros.EOF
            FERubro.AdicionaFila
            FERubro.TextMatrix(i, 1) = rsRubros!cConsDescripcion
            FERubro.TextMatrix(i, 2) = rsRubros!cForCtaCon
            FERubro.TextMatrix(i, 3) = rsRubros!nConsValor
            FERubro.TextMatrix(i, 4) = rsRubros!IdRubroPagoAdelantadoCab
            i = i + 1
            rsRubros.MoveNext
        Loop
    Else
           MsgBox "No existe Rubros registrados", vbInformation, "Aviso"
    End If
    FERubro.lbEditarFlex = False
End Sub

Private Function validarCamposCerrar() As Boolean
Dim i, j As Integer
j = FERubro.Rows

For i = 1 To j - 1
  If FERubro.TextMatrix(i, 2) = "" Then
    MsgBox "La Fórmula del Rubro " & FERubro.TextMatrix(i, 1) & ", esta vaccía. Debe ingresar su plantilla."
    validarCamposCerrar = False
    Exit Function
  End If
Next i

For i = 1 To j - 1
  If Left(FERubro.TextMatrix(i, 2), 2) <> "45" Then
    MsgBox "La Fórmula del Rubro " & FERubro.TextMatrix(i, 1) & ", no inicia con la Cta. Cont. 45. Debe ingresar una plantilla que inicie con la Cta. Cont. 45."
    validarCamposCerrar = False
    Exit Function
  End If
Next i

validarCamposCerrar = True
End Function

Private Function validarCamposGuardar() As Boolean
Dim i As Integer
i = FERubro.Row

  If FERubro.TextMatrix(i, 2) = "" Then
    MsgBox "La Fórmula del Rubro " & FERubro.TextMatrix(i, 1) & ", esta vaccía. Debe ingresar su plantilla."
    validarCamposGuardar = False
    Exit Function
  End If

  If Left(FERubro.TextMatrix(i, 2), 2) <> "45" Then
    MsgBox "La Fórmula del Rubro " & FERubro.TextMatrix(i, 1) & ", no inicia con la Cta. Cont. 45. Debe ingresar una plantilla que inicie con la Cta. Cont. 45."
    validarCamposGuardar = False
    Exit Function
  End If

validarCamposGuardar = True
End Function

Private Sub cmdAceptarRubro_Click()

    Dim oDDocumento As New DDocumento
    Dim oNContFunciones As New NContFunciones
    Dim lsMovNro As String
    Dim lbConfirmarConstante, lbConfirmarRubro As Boolean

If validarCamposGuardar = False Then Exit Sub

If fnAccion = gNuevoRegistro Then
    
    Dim lnCodAnt, lnCodNue As Integer
    
    lbConfirmarConstante = False
    lbConfirmarRubro = False
    
    lsMovNro = oNContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    If Trim(FERubro.TextMatrix(fnFilaNoEditar, 3)) = "" Then
        lnCodAnt = oDDocumento.recuperarCodigoRubro()
        lnCodNue = lnCodAnt + 1
        
        lbConfirmarConstante = oDDocumento.registrarConstanteRubro(lnCodNue, _
                                                                  Trim(FERubro.TextMatrix(fnFilaNoEditar, 1)))
        If lbConfirmarConstante Then
            lbConfirmarRubro = oDDocumento.registrarRubroPagoAdelantado(lnCodNue, _
                                                                       Trim(FERubro.TextMatrix(fnFilaNoEditar, 2)), _
                                                                       lsMovNro)
            If lbConfirmarRubro Then
                MsgBox "Se guardarón correctamente los datos", vbInformation, "Aviso"
                Call CargarFERubro
            Else
                MsgBox "No se guardo la Fórmula de Cta. Cont. del Rubro " & Trim(FERubro.TextMatrix(fnFilaNoEditar, 1)), vbInformation, "Aviso"
            End If
        Else
            MsgBox "No se pudo registrar el Rubro y su Fórmula de Cta. Cont.", vbInformation, "Aviso"
        End If
    End If


End If

If fnAccion = gEditarRegistro Then

   
    lbConfirmarConstante = False
    lbConfirmarRubro = False
    
    lsMovNro = oNContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    lbConfirmarConstante = oDDocumento.modificarConstanteRubro(CInt(FERubro.TextMatrix(fnFilaNoEditar, 3)), _
                                                              Trim(FERubro.TextMatrix(fnFilaNoEditar, 1)))

    If lbConfirmarConstante Then
         If Trim(FERubro.TextMatrix(fnFilaNoEditar, 4)) <> "0" Then
            lbConfirmarRubro = oDDocumento.modificarRubroPagoAdelantado(CInt(FERubro.TextMatrix(fnFilaNoEditar, 4)), _
                                                                       CInt(FERubro.TextMatrix(fnFilaNoEditar, 3)), _
                                                                       Trim(FERubro.TextMatrix(fnFilaNoEditar, 2)), _
                                                                       lsMovNro)
         Else
            lbConfirmarRubro = oDDocumento.registrarRubroPagoAdelantado(CInt(FERubro.TextMatrix(fnFilaNoEditar, 3)), _
                                                                       Trim(FERubro.TextMatrix(fnFilaNoEditar, 2)), _
                                                                       lsMovNro)
         End If
         
        If lbConfirmarRubro Then
            MsgBox "Se modificaron correctamente los datos", vbInformation, "Aviso"
            Call CargarFERubro
        Else
            MsgBox "No se modifico la Fórmula de Cta. Cont. del Rubro " & Trim(FERubro.TextMatrix(fnFilaNoEditar, 1)), vbInformation, "Aviso"
        End If
    Else
        MsgBox "No se pudo modificar el Rubro y su Fórmula de Cta. Cont.", vbInformation, "Aviso"
    End If
    
    Set oNContFunciones = Nothing
    Set oDDocumento = Nothing
    
    cmdNuevoRubro.Visible = True
    cmdEditarRubro.Visible = True
    cmdEliminarRubro.Visible = True
    cmdAceptarRubro.Visible = False
    cmdCancelarRubro.Visible = False
    FERubro.lbEditarFlex = False
End If
fnAccion = gValorDefectoAccion
End Sub

Private Sub cmdCancelarRubro_Click()
Call CargarFERubro
cmdNuevoRubro.Visible = True
cmdEditarRubro.Visible = True
cmdEliminarRubro.Visible = True
cmdAceptarRubro.Visible = False
cmdCancelarRubro.Visible = False
fnAccion = gValorDefectoAccion
End Sub

Private Sub cmdEditarRubro_Click()
cmdNuevoRubro.Visible = False
cmdEditarRubro.Visible = False
cmdEliminarRubro.Visible = False
cmdAceptarRubro.Visible = True
cmdCancelarRubro.Visible = True
fnAccion = gEditarRegistro
FERubro.lbEditarFlex = True
fnFilaNoEditar = FERubro.Row
End Sub

Private Sub cmdEliminarRubro_Click()
Dim oDDocumento As DDocumento
Set oDDocumento = New DDocumento
Dim rsListaMN As ADODB.Recordset
Set rsListaMN = New ADODB.Recordset
Dim rsListaME As ADODB.Recordset
Set rsListaME = New ADODB.Recordset
Dim lbConfirmarConstante, lbConfirmarRubro As Boolean

lbConfirmarConstante = False
lbConfirmarRubro = False

Set rsListaMN = oDDocumento.buscarPagosAdelantados(FERubro.TextMatrix(FERubro.Row, 3), _
                                                  "1", _
                                                 True)
Set rsListaME = oDDocumento.buscarPagosAdelantados(FERubro.TextMatrix(FERubro.Row, 3), _
                                                  "2", _
                                                 True)
If (Not rsListaMN.BOF And Not rsListaMN.EOF) Or (Not rsListaME.BOF And Not rsListaME.EOF) Then
    MsgBox "El Rubro " & Trim(FERubro.TextMatrix(FERubro.Row, 1)) & " no puede ser eleminar porque esta relacionado con algún Pago Adelantado"
    Exit Sub
Else
    If MsgBox("¿Esta seguro que desea eliminar el Rubro " & Trim(FERubro.TextMatrix(FERubro.Row, 1)) & "?", vbYesNo, "Aviso") = vbYes Then
        lbConfirmarRubro = oDDocumento.eliminarRubroPagoAdelantado(CInt(Trim(FERubro.TextMatrix(FERubro.Row, 4))))
        
        If lbConfirmarRubro Then
        lbConfirmarConstante = oDDocumento.eliminarConstanteRubro(CInt(Trim(FERubro.TextMatrix(FERubro.Row, 3))))
            If lbConfirmarConstante Then
                Call CargarFERubro
            End If
        Else
            MsgBox "No pudo elimar el Rubro", vbInformation, "Aviso"
        End If
        
    End If
End If

End Sub


Private Sub Form_Load()
    Call CargarFERubro
End Sub

Private Sub cmdNuevoRubro_Click()
cmdNuevoRubro.Visible = False
cmdEditarRubro.Visible = False
cmdEliminarRubro.Visible = False
cmdAceptarRubro.Visible = True
cmdCancelarRubro.Visible = True
fnAccion = gNuevoRegistro
FERubro.lbEditarFlex = True
FERubro.AdicionaFila
fnFilaNoEditar = FERubro.Rows - 1
End Sub

Private Sub FERubro_Click()
    Call FERubro_RowColChange
End Sub

Private Sub FERubro_EnterCell()
    Call FERubro_RowColChange
End Sub

Private Sub FERubro_RowColChange()
     If FERubro.lbEditarFlex Then
        If fnFilaNoEditar <> -1 Then
           FERubro.Row = fnFilaNoEditar
        End If
    End If
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If validarCamposCerrar = False Then
        Cancel = True
    End If
End Sub

