VERSION 5.00
Begin VB.Form frmCredFormEvalActivosFijos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Activo Fijo :"
   ClientHeight    =   3885
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   9270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCredFormEvalActivosFijos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   1
      Top             =   2880
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
      Top             =   2880
      Width           =   1290
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   3240
      Width           =   9255
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
         Left            =   6720
         TabIndex        =   2
         Top             =   160
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
         Left            =   7920
         TabIndex        =   3
         Top             =   160
         Width           =   1170
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   2655
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   9225
      Begin SICMACT.FlexEdit feActivosFijos 
         Height          =   2175
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   3836
         Cols0           =   6
         HighLight       =   1
         EncabezadosNombres=   "N-Descripcion-Cantidad-Precio-Valor-Aux"
         EncabezadosAnchos=   "350-3700-1500-1700-1700-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         ColumnasAEditar =   "X-1-2-3-4-X"
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-L-R-C-R-C"
         FormatosEdit    =   "0-0-2-2-2-2"
         TextArray0      =   "N"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
      End
   End
   Begin SICMACT.EditMoney txtTotalActivosFijos 
      Height          =   300
      Left            =   7680
      TabIndex        =   7
      Top             =   2940
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      Enabled         =   -1  'True
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
      Left            =   6960
      TabIndex        =   8
      Top             =   2940
      Width           =   525
   End
End
Attribute VB_Name = "frmCredFormEvalActivosFijos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmCredFormEvalActivosFijos
'** Descripción : Formulario para evaluación de Creditos que tienen el tipo de evaluación 2
'** Referencia  : ERS004-2016
'** Creación    : LUCV, 20160525 09:00:00 AM
'**********************************************************************************************
Option Explicit
Dim fvDetalleRef() As tFormEvalDetalleActivosActivoFijoFormato5
Dim fnNroFila As Integer
Dim fnTotal As Double
Dim fvConsCod As Integer
Dim fvConsValor As Integer

Dim sCtaCod As String
Dim nConsCod As Integer
Dim nConsValor As Integer
Dim nTipoPat As Integer
Dim fsTitulo As String

Private Sub Form_Load()
CentraForm Me
Me.Icon = LoadPicture(App.Path & gsRutaIcono)

End Sub

Public Function Inicio(ByRef pvDetalleActivoFlex() As tFormEvalDetalleActivosActivoFijoFormato5, _
                       ByRef pnTotalCeldaFlex As Double, _
                       ByRef pnConsCod As Integer, _
                       ByRef pnConsValor As Integer, ByVal psTitulo As String, _
                       ByVal psCtaCod As String, ByVal pnTipoPat As Integer) As Boolean  'RECO20160719 SE AGREGO LA CUENTA
                       
        
    sCtaCod = psCtaCod
    nConsCod = pnConsCod
    nConsValor = pnConsValor
    nTipoPat = pnTipoPat

    Me.Caption = psTitulo
    fsTitulo = psTitulo
    If UBound(pvDetalleActivoFlex) > 0 Then 'Si Matrix Contiene Datos
        fvDetalleRef = pvDetalleActivoFlex
        Call SetFlexDetalleCtaCobrar
        fnTotal = pnTotalCeldaFlex
        fvConsCod = pnConsCod
        fvConsValor = pnConsValor
        Call SumarMontofeActivosFijos
    Else
        ReDim pvDetalleActivoFlex(CargarDatosDetalle(sCtaCod, nConsCod, nConsValor, nTipoPat))
        'If Not CargarDatosDetalle(sCtaCod, nConsCod, nConsValor, nTipoPat) Then
            'ReDim pvDetalleActivoFlex(0)
            pnTotalCeldaFlex = 0
            fnTotal = 0
        'Else
            Call SumarMontofeActivosFijos
        'End If
    End If
        
    Me.Show 1
    pvDetalleActivoFlex = fvDetalleRef
    pnTotalCeldaFlex = fnTotal
    pnConsCod = fvConsValor
    pnConsValor = fvConsValor
    Inicio = True
    
End Function

Private Sub SetFlexDetalleCtaCobrar()
    Dim Index As Integer
    feActivosFijos.lbEditarFlex = True
    Call LimpiaFlex(feActivosFijos)

    For Index = 1 To UBound(fvDetalleRef)
            feActivosFijos.AdicionaFila
            feActivosFijos.TextMatrix(Index, 1) = fvDetalleRef(Index).cDescripcion
            feActivosFijos.TextMatrix(Index, 2) = fvDetalleRef(Index).nCantidad
            feActivosFijos.TextMatrix(Index, 3) = fvDetalleRef(Index).nPrecio
            feActivosFijos.TextMatrix(Index, 4) = Format(fvDetalleRef(Index).nTotal, "#,##0.00")
    Next
    SumarMontofeActivosFijos
End Sub
Private Sub SumarMontofeActivosFijos()
    Dim i As Integer
    Dim lnCantidad As Currency
    Dim lnPrecio As Currency
    Dim lnTotalFlex As Currency
    Dim lnTotal As Currency
    lnTotal = 0
    If feActivosFijos.TextMatrix(1, 0) <> "" Then
        For i = 1 To feActivosFijos.Rows - 1
            lnCantidad = IIf(IsNumeric(feActivosFijos.TextMatrix(i, 2)), feActivosFijos.TextMatrix(i, 2), 0)
            lnPrecio = IIf(IsNumeric(feActivosFijos.TextMatrix(i, 3)), feActivosFijos.TextMatrix(i, 3), 0)
            lnTotalFlex = CCur(lnCantidad * lnPrecio)
            feActivosFijos.TextMatrix(i, 4) = Format(lnTotalFlex, "#,##0.00")
            lnTotal = lnTotal + lnTotalFlex
        Next
    End If
    txtTotalActivosFijos.Enabled = False
    txtTotalActivosFijos.Text = Format(lnTotal, "#,##0.00")
    fnTotal = Format(lnTotal, "#,##0.00")
End Sub

Private Sub cmdAceptar_Click()
    Dim Index As Integer
    Dim i As Integer
    Dim sMjs As String
    
    sMjs = ValidaDatos
    If sMjs = "" Then
    'If Not validarDetalle Then Exit Sub
        If MsgBox("¿Desea Guardar, " & fsTitulo & "?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        'Llenado de Matriz
        Index = Me.feActivosFijos.Rows - 1
        
        If feActivosFijos.TextMatrix(1, 0) = "" Then
            Index = 0
        End If
        
        ReDim Preserve fvDetalleRef(Index)
            For i = 1 To Index
                fvDetalleRef(i).cDescripcion = feActivosFijos.TextMatrix(i, 1)
                fvDetalleRef(i).nCantidad = feActivosFijos.TextMatrix(i, 2)
                fvDetalleRef(i).nPrecio = feActivosFijos.TextMatrix(i, 3)
                fvDetalleRef(i).nTotal = feActivosFijos.TextMatrix(i, 4)
            Next i
        fnTotal = CDbl(txtTotalActivosFijos.Text)
        Unload Me
        Else
            MsgBox sMjs, vbInformation, "Alerta"
        End If
End Sub

Private Sub cmdAgregar_Click()
    If feActivosFijos.Rows - 1 < 25 Then
        feActivosFijos.lbEditarFlex = True
        feActivosFijos.AdicionaFila
        feActivosFijos.SetFocus
        SendKeys "{Enter}"
    Else
    MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdQuitar_Click()
    If MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        feActivosFijos.EliminaFila (feActivosFijos.row)
        txtTotalActivosFijos.Text = Format(SumarCampo(feActivosFijos, 3), "#,##0.00")
    End If
End Sub

Private Sub cmdCancelar_Click()
'    Dim Index As Integer
'    Dim i As Integer
'
'    If IsArray(fvDetalleRef) Then
'        If UBound(fvDetalleRef) > 0 Then
'        Else
'            Call CargarDatosDetalle(sCtaCod, nConsCod, nConsValor, nTipoPat)
'        End If
'    End If
'    Index = Me.feActivosFijos.Rows - 1
'    ReDim Preserve fvDetalleRef(Index)
'    For i = 1 To Index
'        fvDetalleRef(i).cDescripcion = feActivosFijos.TextMatrix(i, 1)
'        fvDetalleRef(i).nCantidad = CDbl(feActivosFijos.TextMatrix(i, 2))
'        fvDetalleRef(i).nPrecio = feActivosFijos.TextMatrix(i, 3)
'        fvDetalleRef(i).nTotal = feActivosFijos.TextMatrix(i, 4)
'    Next i
'    fnTotal = CDbl(txtTotalActivosFijos.Text)
    Unload Me
End Sub

Private Sub feActivosFijos_OnCellChange(pnRow As Long, pnCol As Long)
    If pnCol = 1 Then
    feActivosFijos.TextMatrix(feActivosFijos.row, 1) = UCase(feActivosFijos.TextMatrix(feActivosFijos.row, 1))
    End If
    Call SumarMontofeActivosFijos
    'txtTotalActivosFijos.Text = Format(SumarCampo(feActivosFijos, 4), "#,##0.00")
End Sub

Private Sub feActivosFijos_OnRowChange(pnRow As Long, pnCol As Long)
feActivosFijos.TextMatrix(feActivosFijos.row, 1) = UCase(feActivosFijos.TextMatrix(feActivosFijos.row, 1))
Call SumarMontofeActivosFijos
End Sub

Public Function CargarDatosDetalle(ByVal psCtaCod As String, ByVal pnConsCod As Integer, ByVal pnConsValor As Integer, ByVal pnTipoPat As Integer) As Integer
    Dim oCred As New COMNCredito.NCOMFormatosEval
    Dim rsDatos As ADODB.Recordset
    Dim nIndice As Integer
    Dim i As Integer
    
    Set rsDatos = oCred.ObtieneDetalleActiFijo(psCtaCod, pnConsCod, pnConsValor, pnTipoPat)
    CargarDatosDetalle = False
    If Not (rsDatos.EOF And rsDatos.BOF) Then
        feActivosFijos.Clear
        FormateaFlex feActivosFijos
        For nIndice = 1 To rsDatos.RecordCount
            feActivosFijos.AdicionaFila
            feActivosFijos.TextMatrix(nIndice, 1) = rsDatos!Descripcion
            feActivosFijos.TextMatrix(nIndice, 2) = rsDatos!Cantidad
            feActivosFijos.TextMatrix(nIndice, 3) = rsDatos!CostoUnitario
            feActivosFijos.TextMatrix(nIndice, 4) = rsDatos!total
            rsDatos.MoveNext
        Next
        CargarDatosDetalle = nIndice - 1
        
        ReDim Preserve fvDetalleRef(CargarDatosDetalle)
        For i = 1 To CargarDatosDetalle
            fvDetalleRef(i).cDescripcion = feActivosFijos.TextMatrix(i, 1)
            fvDetalleRef(i).nCantidad = Trim(Right(feActivosFijos.TextMatrix(i, 2), 10))
            fvDetalleRef(i).nPrecio = feActivosFijos.TextMatrix(i, 3)
            fvDetalleRef(i).nTotal = feActivosFijos.TextMatrix(i, 4)
        Next
    End If
End Function

Private Function ValidaDatos() As String
    Dim nIndice As Integer
    
    ValidaDatos = ""
    If feActivosFijos.TextMatrix(nIndice, 1) <> "" Then
        For nIndice = 1 To feActivosFijos.Rows - 1
            If feActivosFijos.TextMatrix(nIndice, 1) = "" Or feActivosFijos.TextMatrix(nIndice, 2) = "" Or _
               feActivosFijos.TextMatrix(nIndice, 3) = "" Or feActivosFijos.TextMatrix(nIndice, 4) = "" Then
                ValidaDatos = "No se puede grabar datos vacíos"
                Exit Function
            End If
            If val(Replace(feActivosFijos.TextMatrix(nIndice, 2), ",", "")) = 0 Or val(Replace(feActivosFijos.TextMatrix(nIndice, 3), ",", "")) = 0 _
               Or val(Replace(feActivosFijos.TextMatrix(nIndice, 4), ",", "")) = 0 Then
                ValidaDatos = "No se puede grabar valores en 0"
                Exit Function
            End If
        Next
    End If
End Function
