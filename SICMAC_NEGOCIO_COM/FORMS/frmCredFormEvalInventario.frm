VERSION 5.00
Begin VB.Form frmCredFormEvalInventario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inventario"
   ClientHeight    =   3615
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   10110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCredFormEvalInventario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   3000
      Width           =   10095
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7440
         TabIndex        =   3
         Top             =   210
         Width           =   1170
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8760
         TabIndex        =   4
         Top             =   210
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "&Quitar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      TabIndex        =   2
      Top             =   2640
      Width           =   1170
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "A&gregar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   1290
   End
   Begin VB.Frame Frame3 
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
      Height          =   2535
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10065
      Begin SICMACT.FlexEdit feInventario 
         Height          =   2175
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   9855
         _extentx        =   17383
         _extenty        =   3836
         cols0           =   7
         highlight       =   1
         encabezadosnombres=   "N-Mercaderia-Cantidad-Unid. Medida-Costo Unitario-Total-Aux"
         encabezadosanchos=   "350-3000-1500-1700-1500-1700-0"
         font            =   "frmCredFormEvalInventario.frx":030A
         font            =   "frmCredFormEvalInventario.frx":0332
         font            =   "frmCredFormEvalInventario.frx":035A
         font            =   "frmCredFormEvalInventario.frx":0382
         font            =   "frmCredFormEvalInventario.frx":03AA
         fontfixed       =   "frmCredFormEvalInventario.frx":03D2
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1  'True
         tipobusqueda    =   3
         columnasaeditar =   "X-1-2-3-4-X-X"
         listacontroles  =   "0-0-0-3-0-0-0"
         encabezadosalineacion=   "L-L-R-L-R-R-L"
         formatosedit    =   "0-0-2-0-2-2-0"
         textarray0      =   "N"
         lbeditarflex    =   -1  'True
         lbbuscaduplicadotext=   -1  'True
         colwidth0       =   345
         rowheight0      =   300
      End
   End
   Begin SICMACT.EditMoney txtTotalInventario 
      Height          =   300
      Left            =   8400
      TabIndex        =   6
      Top             =   2700
      Width           =   1500
      _extentx        =   2646
      _extenty        =   529
      font            =   "frmCredFormEvalInventario.frx":03F8
      text            =   "0"
      enabled         =   -1  'True
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7800
      TabIndex        =   7
      Top             =   2700
      Width           =   525
   End
End
Attribute VB_Name = "frmCredFormEvalInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmCredFormEvalInventario
'** Descripción : Formulario para evaluación de Creditos que tienen el tipo de evaluación 2
'** Referencia  : ERS004-2016
'** Creación    : LUCV, 20160525 09:00:00 AM
'**********************************************************************************************
Option Explicit
Dim fvDetalleRef() As tFormEvalDetalleActivosInventarioFormato5
Dim fnNroFila As Integer
Dim fnTotal As Double
Dim fvConsCod As Integer
Dim fvConsValor As Integer

Dim sCtaCod As String
Dim nConsCod As Integer
Dim nConsValor As Integer
Dim nTipoPat As Integer
Dim vMatListaUnidMed As Variant
Dim fsTitulo As String

Private Sub Form_Load()
    CentraForm Me
    Call CargarControles
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
End Sub

Public Function Inicio(ByRef pvDetalleActivoFlex() As tFormEvalDetalleActivosInventarioFormato5, _
                       ByRef pnTotalCeldaFlex As Double, _
                       ByRef pnConsCod As Integer, _
                       ByRef pnConsValor As Integer, ByVal psTitulo As String, _
                       ByVal psCtaCod As String, ByVal pnTipoPat As Integer) As Boolean  'RECO20160719 SE AGREGO LA CUENTA
    Dim oCons As New COMDConstantes.DCOMConstantes
    feInventario.CargaCombo oCons.RecuperaConstantes(7035)
    
    sCtaCod = psCtaCod
    nConsCod = pnConsCod
    nConsValor = pnConsValor
    nTipoPat = pnTipoPat
    
    Me.Caption = psTitulo
    fsTitulo = psTitulo
    If UBound(pvDetalleActivoFlex) > 0 Then 'Si Matrix Contiene Datos
        fvDetalleRef = pvDetalleActivoFlex
        Call SetFlexDetalleInventario
        fnTotal = pnTotalCeldaFlex
        fvConsCod = pnConsCod
        fvConsValor = pnConsValor
        Call SumarMontofeInventario
    Else
        ReDim pvDetalleActivoFlex(CargarDatosDetalle(psCtaCod, pnConsCod, pnConsValor, pnTipoPat))
        pnTotalCeldaFlex = 0
        fnTotal = 0
        Call SumarMontofeInventario
    End If
        
    Me.Show 1
    pvDetalleActivoFlex = fvDetalleRef
    pnTotalCeldaFlex = fnTotal
    pnConsCod = fvConsValor
    pnConsValor = fvConsValor
    Inicio = True
    
End Function

Private Sub SetFlexDetalleInventario()
    Dim index As Integer, j As Integer
    feInventario.lbEditarFlex = True
    Call LimpiaFlex(feInventario)

    For index = 1 To UBound(fvDetalleRef)
            feInventario.AdicionaFila
            feInventario.TextMatrix(index, 1) = fvDetalleRef(index).cMercaderia
            feInventario.TextMatrix(index, 2) = fvDetalleRef(index).nCantidad
            'feInventario.TextMatrix(Index, 3) = fvDetalleRef(Index).cUnidMed
            For j = 0 To UBound(vMatListaUnidMed) - 1
                If Trim(Right(vMatListaUnidMed(j), 10)) = fvDetalleRef(index).cUnidMed Then
                    feInventario.TextMatrix(index, 3) = vMatListaUnidMed(j)
                    Exit For
                End If
            Next
            feInventario.TextMatrix(index, 4) = Format(fvDetalleRef(index).nCostoUnit, "#,##0.00")
            feInventario.TextMatrix(index, 5) = Format(fvDetalleRef(index).nTotal, "#,##0.00")
    Next
    SumarMontofeInventario
End Sub
Private Sub SumarMontofeInventario()
    Dim i As Integer
    Dim lnMonto As Currency
    Dim lnTotal As Currency
    lnTotal = 0
    If feInventario.TextMatrix(1, 0) <> "" Then
        For i = 1 To feInventario.Rows - 1
            lnMonto = IIf(IsNumeric(feInventario.TextMatrix(i, 5)), feInventario.TextMatrix(i, 5), 0)
            lnTotal = lnTotal + lnMonto
        Next
    End If
    txtTotalInventario.Enabled = False
    txtTotalInventario.Text = Format(lnTotal, "#,##0.00")
    fnTotal = Format(lnTotal, "#,##0.00")
End Sub

Private Sub SumarMontos()
    Dim i As Integer
    Dim lnCantidad As Double
    Dim lnMonto As Currency
    Dim lnTotal As Currency
    lnTotal = 0
    If feInventario.TextMatrix(1, 0) <> "" Then
        For i = 1 To feInventario.Rows - 1
            lnCantidad = IIf(IsNumeric(feInventario.TextMatrix(i, 2)), feInventario.TextMatrix(i, 2), 0)
            lnMonto = IIf(IsNumeric(feInventario.TextMatrix(i, 4)), feInventario.TextMatrix(i, 4), 0)
            lnTotal = CCur(lnCantidad * lnMonto)
            feInventario.TextMatrix(i, 5) = Format(lnTotal, "#,##0.00")
        Next
    End If
End Sub

Private Sub cmdAceptar_Click()
    Dim index As Integer
    Dim i As Integer
    Dim sMsj As String
    
    sMsj = ValidaCamposVacios
    
    If sMsj <> "" Then
        MsgBox sMsj, vbInformation, "Alerta"
        Exit Sub
    End If
    If MsgBox("Desea Guardar, " & fsTitulo & "?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    
    index = Me.feInventario.Rows - 1
    If feInventario.TextMatrix(0, 1) = "" Then
        index = 0
    End If
    
    ReDim Preserve fvDetalleRef(index)
    For i = 1 To index
        fvDetalleRef(i).cMercaderia = feInventario.TextMatrix(i, 1)
        fvDetalleRef(i).nCantidad = feInventario.TextMatrix(i, 2)
        fvDetalleRef(i).cUnidMed = Trim(Right(feInventario.TextMatrix(i, 3), 3))
        fvDetalleRef(i).nCostoUnit = feInventario.TextMatrix(i, 4)
        fvDetalleRef(i).nTotal = feInventario.TextMatrix(i, 5)
    Next i
    fnTotal = CDbl(txtTotalInventario.Text)
    Unload Me
End Sub

Private Sub cmdAgregar_Click()
    If feInventario.Rows - 1 < 25 Then
        feInventario.lbEditarFlex = True
        feInventario.AdicionaFila
        feInventario.SetFocus
        SendKeys "{Enter}"
    Else
    MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdQuitar_Click()
    If MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        feInventario.EliminaFila (feInventario.row)
        txtTotalInventario.Text = Format(SumarCampo(feInventario, 5), "#,##0.00")
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub feInventario_OnCellChange(pnRow As Long, pnCol As Long)
    Select Case pnCol
        Case 1
            feInventario.TextMatrix(feInventario.row, 1) = UCase(feInventario.TextMatrix(feInventario.row, 1))
        Case 2, 4
            If Not IsNumeric(feInventario.TextMatrix(pnRow, pnCol)) Then
                feInventario.TextMatrix(pnRow, pnCol) = "0.00"
            Else
                If feInventario.TextMatrix(pnRow, pnCol) < 0 Then
                    feInventario.TextMatrix(pnRow, pnCol) = "0.00"
                End If
            End If
    End Select
    Call SumarMontos
    txtTotalInventario.Text = Format(SumarCampo(feInventario, 5), "#,##0.00")
End Sub

Private Sub feInventario_OnRowChange(pnRow As Long, pnCol As Long)
feInventario.TextMatrix(feInventario.row, 1) = UCase(feInventario.TextMatrix(feInventario.row, 1))
End Sub

Public Function CargarDatosDetalle(ByVal psCtaCod As String, ByVal pnConsCod As Integer, ByVal pnConsValor As Integer, ByVal pnTipoPat As Integer) As Integer
    Dim oCred As New COMNCredito.NCOMFormatosEval
    Dim rsDatos As ADODB.Recordset
    Dim nIndice As Integer
    Dim i As Integer
    
    Set rsDatos = oCred.ObtieneDetalleInventario(psCtaCod, pnConsCod, pnConsValor, pnTipoPat)
    
    If Not (rsDatos.EOF And rsDatos.BOF) Then
        feInventario.Clear
        FormateaFlex feInventario
        For nIndice = 1 To rsDatos.RecordCount
            feInventario.AdicionaFila
            feInventario.TextMatrix(nIndice, 1) = rsDatos!Mercaderia
            feInventario.TextMatrix(nIndice, 2) = rsDatos!Cantidad
            feInventario.TextMatrix(nIndice, 3) = rsDatos!UnidadMedida
            feInventario.TextMatrix(nIndice, 4) = rsDatos!CostoUnitario
            feInventario.TextMatrix(nIndice, 5) = rsDatos!Total
            rsDatos.MoveNext
        Next
        CargarDatosDetalle = nIndice - 1
        
        ReDim Preserve fvDetalleRef(CargarDatosDetalle)
        For i = 1 To CargarDatosDetalle
            fvDetalleRef(i).cMercaderia = feInventario.TextMatrix(i, 1)
            fvDetalleRef(i).nCantidad = feInventario.TextMatrix(i, 2)
            fvDetalleRef(i).cUnidMed = Trim(Right(feInventario.TextMatrix(i, 3), 3))
            fvDetalleRef(i).nCostoUnit = feInventario.TextMatrix(i, 4)
            fvDetalleRef(i).nTotal = feInventario.TextMatrix(i, 5)
        Next i
    End If
End Function

Private Sub CargarControles()
    Dim oCons As New COMDConstantes.DCOMConstantes
    Dim oDCred As New COMDCredito.DCOMFormatosEval
    Dim rsLista As New ADODB.Recordset
    Dim i As Integer
    
    Set rsLista = oCons.RecuperaConstantes(7035)
    
    If Not (rsLista.EOF And rsLista.BOF) Then
        ReDim vMatListaUnidMed(rsLista.RecordCount)
        
        For i = 1 To rsLista.RecordCount
            vMatListaUnidMed(i - 1) = rsLista!cConsDescripcion & Space(75) & rsLista!nConsValor
            rsLista.MoveNext
        Next
    End If
    Set oDCred = Nothing
End Sub

Private Function ValidaCamposVacios() As String
    Dim nIndice As Integer
    ValidaCamposVacios = ""
    If feInventario.TextMatrix(0, 1) <> "" Then
        For nIndice = 1 To feInventario.Rows - 1
            If feInventario.TextMatrix(nIndice, 1) = "" Or feInventario.TextMatrix(nIndice, 2) = "" Or feInventario.TextMatrix(nIndice, 3) = "" Or feInventario.TextMatrix(nIndice, 4) = "" Or feInventario.TextMatrix(nIndice, 5) = "" Then
                ValidaCamposVacios = "No se puede registrar valores vacios"
            Else
                If val(Replace(feInventario.TextMatrix(nIndice, 2), ",", "")) = 0 Or val(Replace(feInventario.TextMatrix(nIndice, 4), ",", "")) = 0 Or val(Replace(feInventario.TextMatrix(nIndice, 5), ",", "")) = 0 Then
                    ValidaCamposVacios = "Debe ingresar valores mayores a 0"
                End If
            End If
        Next
    End If
End Function
