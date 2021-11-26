VERSION 5.00
Begin VB.Form frmCredFormEvalIfisDetalleFormato6 
   Caption         =   "Ctas x Cobrar"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9900
   ClipControls    =   0   'False
   Icon            =   "frmCredFormEvalIfisDetalleFormato6.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   9900
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDetalle 
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
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   9735
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Quitar"
         Height          =   350
         Left            =   2010
         TabIndex        =   4
         Top             =   3420
         Width           =   900
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   350
         Left            =   2010
         TabIndex        =   7
         Top             =   3420
         Visible         =   0   'False
         Width           =   900
      End
      Begin SICMACT.FlexEdit feDetalle 
         Height          =   3135
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   5530
         Cols0           =   6
         HighLight       =   1
         EncabezadosNombres=   "#-Fecha-Descripción-Monto-nConsCod-nConsValor"
         EncabezadosAnchos=   "400-1000-6800-1200-0-0"
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
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-2-3-X-X"
         ListaControles  =   "0-2-3-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-R-L-L"
         FormatosEdit    =   "0-0-0-2-0-0"
         CantEntero      =   12
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "A&gregar"
         Height          =   350
         Left            =   1065
         TabIndex        =   0
         Top             =   3420
         Width           =   900
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Aceptar"
         Height          =   350
         Left            =   1065
         TabIndex        =   6
         Top             =   3420
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label txtMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   8160
         TabIndex        =   9
         Top             =   3420
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   7680
         TabIndex        =   8
         Top             =   3480
         Width           =   405
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5145
      TabIndex        =   3
      ToolTipText     =   "Cancelar"
      Top             =   3960
      Width           =   1000
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      ToolTipText     =   "Aceptar"
      Top             =   3960
      Width           =   1000
   End
End
Attribute VB_Name = "frmCredFormEvalIfisDetalleFormato6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************************************
'** Nombre : frmGarantiaValorDirectoDetallado
'** Descripción : Para registro/edición/consulta de Valorización Directa Detallada creado segun TI-ERS063-2014
'** Creación : EJVG, 20130115 04:17:43 PM
'**
'** FORMULARIO REUTILIZADO
'** Nombre : frmCredFormEvalDetalleFormato6
'** Descripción : Para registro/edición/consulta detalles de Estados financieros en Formatos de evaluacion
'** Modificación : PEAC, 20160618 12:32:01 PM
'*************************************************************************************************************
Option Explicit

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&

Dim fbRegistrar As Boolean
Dim fbEditar As Boolean
Dim fbConsultar As Boolean

Dim fbOk As Boolean
Dim fbPrimero As Boolean
Dim fnMoneda As Moneda
Dim fsGlosa As String

'LUCV20170915 *****-> Comentó y agregó, según ERS051-2017
'Dim fvDetalle() As tForEvalEstFinFormato6
'Dim fvDetalle_ULT() As tForEvalEstFinFormato6
'Detalle Activos y Pasivos
Dim fvDetalle() As tFormEvalDetalleEstFinFormato6
Dim fvDetalle_ULT() As tFormEvalDetalleEstFinFormato6 'Matriz Detalle
'<***** Fin LUCV20170915
Dim fnNoMoverFila As Integer
Dim fbFocoGrilla As Boolean
Dim fRsDocs As ADODB.Recordset
Dim fnTotal As Double
Dim fnTotalOriginal As Double
Dim fvConsCod As Integer
Dim fvConsValor As Integer
Dim fsFechaReg As String

Private Function ValidaIfiExisteDuplicadoLista(ByVal psCodIfi As String, ByVal pnFila As Integer) As Boolean
    Dim I As Integer
    
    ValidaIfiExisteDuplicadoLista = False
    
    For I = 1 To Me.feDetalle.rows - 1
        If Trim(Right(feDetalle.TextMatrix(I, 1), 13)) = psCodIfi Then 'LUCV20161115, Modificó->Según ERS068-2016(10-13)
            If I <> pnFila Then
                ValidaIfiExisteDuplicadoLista = True
                Exit Function
            End If
        End If
    Next
End Function


Public Function DisableCloseButton(frm As Form) As Boolean
'PURPOSE: Removes X button from a form
'EXAMPLE: DisableCloseButton Me
'RETURNS: True if successful, false otherwise
'NOTES:   Also removes Exit Item from
'         Control Box Menu
    Dim lHndSysMenu As Long
    Dim lAns1 As Long, lAns2 As Long
    
    lHndSysMenu = GetSystemMenu(frm.hwnd, 0)

    'remove close button
    lAns1 = RemoveMenu(lHndSysMenu, 6, MF_BYPOSITION)

   'Remove seperator bar
    lAns2 = RemoveMenu(lHndSysMenu, 5, MF_BYPOSITION)
    
    'Return True if both calls were successful
    DisableCloseButton = (lAns1 <> 0 And lAns2 <> 0)

End Function

Public Function Registrar(ByVal pbPrimero As Boolean, ByRef pnMoneda As Moneda, ByRef psGlosa As String, ByRef pvDetalle() As tFormEvalDetalleEstFinFormato6, ByRef pvDetalle_ULT() As tFormEvalDetalleEstFinFormato6, pnTotal As Double, pnConsCod As Integer, pnConsValor As Integer, psFechaReg As String) As Boolean
    fbRegistrar = True
    fbPrimero = pbPrimero
    fnMoneda = pnMoneda
    fsGlosa = psGlosa
    fnTotal = pnTotal
    fnTotalOriginal = pnTotal
    fvConsCod = pnConsCod
    fvConsValor = pnConsValor
    fvDetalle = pvDetalle
    fvDetalle_ULT = pvDetalle_ULT
    fsFechaReg = psFechaReg
    Show 1
    pnMoneda = fnMoneda
    psGlosa = fsGlosa
    pvDetalle = fvDetalle
    pnTotal = fnTotal
    pnConsCod = fvConsValor
    pnConsValor = fvConsValor
    Registrar = fbOk
End Function
Public Function Editar(ByVal pbPrimero As Boolean, ByRef pnMoneda As Moneda, ByRef psGlosa As String, ByRef pvDetalle() As tFormEvalDetalleEstFinFormato6, ByRef pvDetalle_ULT() As tFormEvalDetalleEstFinFormato6) As Boolean
    fbEditar = True
    fnMoneda = pnMoneda
    fsGlosa = psGlosa
    fbPrimero = pbPrimero
    fvDetalle = pvDetalle
    fvDetalle_ULT = pvDetalle_ULT
    Show 1
    pnMoneda = fnMoneda
    psGlosa = fsGlosa
    pvDetalle = fvDetalle
    
    Editar = fbOk
End Function
Public Sub Consultar(ByVal pnMoneda As Moneda, ByVal pnMonto As Currency, ByVal psGlosa As String, ByRef pvDetalle() As tFormEvalDetalleEstFinFormato6)
    fnMoneda = pnMoneda
    fsGlosa = psGlosa
    fvDetalle = pvDetalle
    fbConsultar = True
    Show 1
End Sub
Private Sub cmdCancel_Click()
    SetFlexDetalle
    EditarDetalle False
    fnNoMoverFila = -1
    
    cmdAceptar.Enabled = True
End Sub
Private Sub cmdOK_Click()
    Dim index As Integer
    Dim lsDocTpo As String
    Dim I As Integer
    Dim j As Integer
    
    If Not validarDetalle Then Exit Sub
    
    For I = 1 To Me.feDetalle.rows - 1
        For j = I + 1 To Me.feDetalle.rows - 1
            If Me.feDetalle.TextMatrix(I, 1) = Me.feDetalle.TextMatrix(j, 1) And _
               Me.feDetalle.TextMatrix(I, 2) = Me.feDetalle.TextMatrix(j, 2) And _
               Me.feDetalle.TextMatrix(I, 3) = Me.feDetalle.TextMatrix(j, 3) _
            Then
                MsgBox "Este registro está duplicado, por favor modifique..."
                Exit Sub
            End If
        Next j
    Next I
    
    For I = 1 To Me.feDetalle.rows - 1
        For j = I + 1 To Me.feDetalle.rows - 1
            If Me.feDetalle.TextMatrix(I, 2) = Me.feDetalle.TextMatrix(j, 2) Then
                MsgBox "No se puede registrar dos veces una misma IFI..."
                Exit Sub
            End If
        Next j
    Next I
    
    index = UBound(fvDetalle) + 1
    ReDim Preserve fvDetalle(index)
    
    fvDetalle(index).dFecha = feDetalle.TextMatrix(fnNoMoverFila, 1)
    fvDetalle(index).cDescripcion = feDetalle.TextMatrix(fnNoMoverFila, 2)
    fvDetalle(index).nImporte = feDetalle.TextMatrix(fnNoMoverFila, 3)

    SetFlexDetalle
    EditarDetalle False
    fnNoMoverFila = -1
    
    cmdAceptar.Enabled = True
End Sub

Private Sub feDetalle_GotFocus()
    fbFocoGrilla = True
End Sub
Private Sub feDetalle_LostFocus()
    fbFocoGrilla = False
End Sub

Private Sub feDetalle_OnCellChange(pnRow As Long, pnCol As Long)
    If pnCol = 2 Or pnCol = 3 Then
        SumarizarDetalle
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If fbFocoGrilla Then
        If KeyCode = 86 And Shift = 2 Then
            KeyCode = 10
        End If
    End If
End Sub
Private Sub Form_Load()
    
    DisableCloseButton Me
    
    fbOk = False
    Screen.MousePointer = 11
    
    CargarControles
    CargarVariables
    LimpiarControles

    If fbEditar Or fbConsultar Then
        SetFlexDetalle
        
        If fbConsultar Then
            fraDetalle.Enabled = False
            cmdAceptar.Enabled = False
        End If
    End If
    
    If fbPrimero Then
        fnMoneda = gMonedaNacional
    End If
    
    If UBound(fvDetalle) > 0 Then
    
    Else
        ReDim fvDetalle_ULT(0)
        ReDim fvDetalle(0)
    End If
    
    If fbRegistrar Then
        If UBound(fvDetalle_ULT) > 0 Then
            fvDetalle = fvDetalle_ULT
        End If
        SetFlexDetalle
    End If
       
    If fbRegistrar Then
        Caption = fsGlosa & " [ NUEVO ]"
    End If
    If fbConsultar Then
        Caption = fsGlosa & " [ CONSULTAR ]"
    End If
    If fbEditar Then
        Caption = fsGlosa & " [ EDITAR ]"
    End If
    
    Screen.MousePointer = 0
End Sub
Private Sub CargarVariables()
    Dim oCons As New COMDConstantes.DCOMConstantes
    Set fRsDocs = oCons.RecuperaConstantes(gColocPigTipoDocumento)
    Set oCons = Nothing
End Sub
Private Sub CargarControles()
    Dim oCons As New COMDConstantes.DCOMConstantes
    Dim rsMoneda As New ADODB.Recordset

    Dim oDCred As New COMDCredito.DCOMFormatosEval
    Dim rsLista As New ADODB.Recordset
    Dim rsLista2 As New ADODB.Recordset
    Dim I As Integer
    
    Set rsLista = oDCred.CargarIfisDetalleForm6(fvConsCod, fvConsValor) 'LUCV20161115, Modificó->Según ERS068-2016 (Agregó parametros)
'    Set rsLista2 = oDCred.CargarOtrasIfis()
    
    feDetalle.CargaCombo rsLista

    RSClose rsMoneda
    Set oCons = Nothing
End Sub
Private Sub LimpiarControles()
    txtMonto.Caption = "0.00"
    FormateaFlex feDetalle
End Sub

Private Sub cmdCancelar_Click()
    fnTotal = fnTotalOriginal
    fbOk = False
    Unload Me
End Sub
Private Sub cmdAceptar_Click()
    If Not validarDetalle Then Exit Sub
    
    fnTotal = txtMonto.Caption
        
    fbOk = True
    Unload Me
End Sub
Private Sub cmdEliminar_Click()
    Dim lvDetalleTmp() As tFormEvalDetalleEstFinFormato6
    Dim index As Integer, indexTmp As Integer
    
    On Error GoTo ErrEliminar
    If feDetalle.TextMatrix(1, 0) = "" Then Exit Sub
    
    ReDim lvDetalleTmp(0)
    For index = 1 To UBound(fvDetalle)
        ReDim Preserve lvDetalleTmp(index)
        lvDetalleTmp(index) = fvDetalle(index)
    Next
    
    index = 0
    ReDim fvDetalle(0)
    For indexTmp = 1 To UBound(lvDetalleTmp)
        If indexTmp <> feDetalle.row Then
            index = index + 1
            ReDim Preserve fvDetalle(index)
            fvDetalle(index) = lvDetalleTmp(indexTmp)
        End If
    Next
    Erase lvDetalleTmp
    
    SetFlexDetalle
    Exit Sub
ErrEliminar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdAgregar_Click()
    If feDetalle.TextMatrix(1, 0) <> "" Then
        If Not validarDetalle Then Exit Sub
    End If
    
    feDetalle.AdicionaFila
        
    feDetalle.SetFocus
    SendKeys "{ENTER}"
    
    EditarDetalle True
    fnNoMoverFila = feDetalle.row
    
    Me.feDetalle.TextMatrix(feDetalle.row, 1) = fsFechaReg
    
    cmdAceptar.Enabled = False
End Sub
Public Function validarDetalle() As Boolean
    Dim I As Integer, j As Integer
    
    If feDetalle.TextMatrix(1, 0) = "" Then
        If MsgBox("No hay datos para Grabar, desea continuar?", vbQuestion + vbYesNo, "Pregunta") = vbNo Then
            EnfocaControl cmdAgregar
            Exit Function
        Else
            validarDetalle = True
            Exit Function
        End If

'        If MsgBox("No hay datos para Grabar, desea continuar?", vbQuestion + vbYesNo, "Pregunta") = vbNo Then
''            MsgBox "Ud. debe ingresar el Detalle de la Valorización", vbInformation, "Aviso"
'            EnfocaControl cmdAgregar
'            Exit Function
'        End If
    End If
    For I = 1 To feDetalle.rows - 1
        For j = 1 To feDetalle.cols - 1
            If feDetalle.ColWidth(j) > 0 Then
                If Len(Trim(feDetalle.TextMatrix(I, j))) = 0 Then
                    MsgBox "El campo " & UCase(feDetalle.TextMatrix(0, j)) & " está vacio, verifique..", vbInformation, "Aviso"
                    EnfocaControl feDetalle
                    feDetalle.TopRow = I
                    feDetalle.row = I
                    feDetalle.Col = j
                    Exit Function
                End If
            End If
        Next
    Next
    For I = 1 To feDetalle.rows - 1
        For j = 1 To feDetalle.cols - 1
            If j = 3 Then 'Valida se hayan ingresado Cantidad
                If val(feDetalle.TextMatrix(I, j)) <= 0 Then
                    MsgBox "Ud. debe ingresar un valor mayor a cero", vbInformation, "Aviso"
                    EnfocaControl feDetalle
                    feDetalle.TopRow = I
                    feDetalle.row = I
                    feDetalle.Col = j
                    Exit Function
                End If
            End If
        Next
    Next
    validarDetalle = True
End Function
Private Sub SetFlexDetalle()
    Dim index As Integer
    Dim vResultado() As tFormEvalDetalleEstFinFormato6
    
    FormateaFlex feDetalle

    For index = 1 To UBound(fvDetalle)
            feDetalle.AdicionaFila
            feDetalle.TextMatrix(index, 1) = fvDetalle(index).dFecha
            feDetalle.TextMatrix(index, 2) = fvDetalle(index).cDescripcion
            feDetalle.TextMatrix(index, 3) = Format(fvDetalle(index).nImporte, "#,##0.00")
    Next
    
    SumarizarDetalle
End Sub
Private Sub EditarDetalle(ByVal pbEditar As Boolean)
    cmdAgregar.Visible = Not pbEditar
    cmdEliminar.Visible = Not pbEditar
    cmdOK.Visible = pbEditar
    cmdCancel.Visible = pbEditar
    
    feDetalle.lbEditarFlex = pbEditar
End Sub
Private Sub feDetalle_RowColChange()
    Dim rs As ADODB.Recordset
    If feDetalle.lbEditarFlex Then
        If feDetalle.Col = 4 Then
            Set rs = fRsDocs.Clone
            feDetalle.CargaCombo rs
        End If
        feDetalle.row = fnNoMoverFila
    End If
    feDetalle.TextMatrix(feDetalle.row, 1) = UCase(feDetalle.TextMatrix(feDetalle.row, 1))
    RSClose rs
End Sub

Private Sub SumarizarDetalle()
    Dim I As Integer
    Dim lnCantidad As Integer
    Dim lnPrecioUnit As Currency
    Dim lnTotal As Currency
    Dim lnMonto As Currency
        
    If feDetalle.TextMatrix(1, 0) <> "" Then
        For I = 1 To feDetalle.rows - 1
            lnMonto = IIf(IsNumeric(feDetalle.TextMatrix(I, 3)), feDetalle.TextMatrix(I, 3), 0)
            lnTotal = lnTotal + lnMonto
        Next
    End If
    
    txtMonto.Caption = Format(lnTotal, "#,##0.00")
End Sub

