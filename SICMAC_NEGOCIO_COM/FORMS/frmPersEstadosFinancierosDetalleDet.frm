VERSION 5.00
Begin VB.Form frmPersEstadosFinancierosDetalleDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Titulo"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9840
   Icon            =   "frmPersEstadosFinancierosDetalleDet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameGuardarCancelar 
      Height          =   615
      Left            =   5140
      TabIndex        =   2
      Top             =   4320
      Width           =   4695
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
         Left            =   2400
         TabIndex        =   4
         Top             =   180
         Width           =   2130
      End
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
         Left            =   270
         TabIndex        =   3
         Top             =   180
         Width           =   1890
      End
   End
   Begin VB.Frame frameDetalle 
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
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   9700
      Begin VB.Frame frameBotonesDetalle 
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   3000
         Width           =   4695
         Begin VB.CommandButton cmdQuitarDetalle 
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
            Left            =   2355
            TabIndex        =   9
            Top             =   160
            Width           =   2140
         End
         Begin VB.CommandButton cmdAgregarDetalle 
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
            TabIndex        =   8
            Top             =   160
            Width           =   2140
         End
      End
      Begin SICMACT.FlexEdit feDetalle 
         Height          =   2775
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   4895
         Cols0           =   4
         HighLight       =   1
         EncabezadosNombres=   "#-Detalle del ítem-Monto-aux"
         EncabezadosAnchos=   "500-7000-1800-0"
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
         ColumnasAEditar =   "X-1-2-X"
         ListaControles  =   "0-3-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-L-R-C"
         FormatosEdit    =   "0-0-2-2"
         CantEntero      =   12
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
      End
      Begin SICMACT.EditMoney txtTotalDetalle 
         Height          =   300
         Left            =   7200
         TabIndex        =   5
         Top             =   3240
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Label lblTotalDetalle 
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
         Left            =   6600
         TabIndex        =   6
         Top             =   3285
         Width           =   525
      End
   End
   Begin VB.Label Label2 
      Caption         =   "(Expresado en soles)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8040
      TabIndex        =   11
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Ingresar detalle y monto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   105
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   9480
      Y1              =   360
      Y2              =   360
   End
End
Attribute VB_Name = "frmPersEstadosFinancierosDetalleDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmPersEstadosFinancierosDetalleDet
'** Descripción : Formulario para detalle de los estados financieros (Módulo Persona)
'** Referencia  : ERS051-2017
'** Creación    : LUCV, 20170915 09:00:00 AM
'**********************************************************************************************
Option Explicit
    Dim fvDetalleRef() As tfrmPersEstadosFinancierosDetalle
    Dim fnNroFila As Integer
    Dim fnTotal As Double
    Dim fvConsCod As Integer
    Dim fvConsValor As Integer
    Dim fvConsValorGrupo As Integer
    Dim vMatrizListaItem As Variant
    Dim vMatListaIfis As Variant
    Dim fsTitulo As String
    Dim nTotalSubDetalle As Currency

Private Sub Form_Activate()
    EnfocaControl cmdAgregarDetalle
End Sub

Private Sub Form_Load()
    CentraForm Me
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
End Sub
Public Function Inicio(ByRef pvDetalleFlex() As tfrmPersEstadosFinancierosDetalle, _
                       ByRef pnTotalCeldaFlex As Double, _
                       ByRef pnConsCod As Integer, _
                       ByRef pnConsValor As Integer, _
                       ByVal psTitulo As String, _
                       ByVal pnCodEF As Integer, _
                       Optional pnConsValorGrupo As Integer = -1) As Boolean
    Me.Caption = psTitulo
    fsTitulo = psTitulo
    fvConsCod = pnConsCod
    fvConsValor = pnConsValor
    fvConsValorGrupo = pnConsValorGrupo

    If fvConsValor = 45 Then 'Modifica Tamaño si NO tiene Ifis
        Call CargarControlesDetalleDet(fvConsCod, fvConsValor)
    Else
        Call CargarControlesDetalle(fvConsCod, fvConsValor)
    End If
    
    If IsArray(pvDetalleFlex) Then
        If UBound(pvDetalleFlex) > 0 Then 'Si Matrix Contiene Datos
            fvDetalleRef = pvDetalleFlex
            If fvConsValor = 45 Then
                Call CargarFlexConArrayDetalleDet
            Else
                Call CargarFlexConArrayDetalle
            End If
            
            fnTotal = pnTotalCeldaFlex
            Call SumarMontofeDetalle
        Else
            ReDim pvDetalleFlex(CargaDatosMantenimiento(pnCodEF, pnConsCod, pnConsValor, pnConsValorGrupo))
            Call SumarMontofeDetalle
        End If
    Else
        ReDim pvDetalleFlex(0)
        pnTotalCeldaFlex = 0
        fnTotal = 0
    End If
       
    Me.Show 1
    pvDetalleFlex = fvDetalleRef
    pnTotalCeldaFlex = fnTotal
    pnConsCod = fvConsValor
    pnConsValor = fvConsValor
    Inicio = True
End Function
Private Sub CargarControlesDetalle(Optional pnConsCod As Integer = -1, Optional pnConsValor As Integer = -1)
    Dim oDCred As New COMDCredito.DCOMFormatosEval
    Dim rsListasItems As New ADODB.Recordset
    Dim rsListasItems2 As New ADODB.Recordset
    Dim i As Integer
    
    Set rsListasItems = oDCred.CargarDetalleItemEEFF(pnConsCod, pnConsValor)
    Set rsListasItems2 = oDCred.CargarDetalleItemEEFF(pnConsCod, pnConsValor)
    
    feDetalle.CargaCombo rsListasItems
    If Not (rsListasItems2.EOF And rsListasItems2.BOF) Then
        ReDim vMatrizListaItem(rsListasItems2.RecordCount)
        For i = 1 To rsListasItems2.RecordCount
            vMatrizListaItem(i - 1) = rsListasItems2!cSubItem
            rsListasItems2.MoveNext
        Next
    End If
    Set oDCred = Nothing
End Sub
Private Sub CargarControlesDetalleDet(Optional pnConsCod As Integer = -1, Optional pnConsValor As Integer = -1)
    Dim oDCred As New COMDCredito.DCOMFormatosEval
    Dim rsListaItemsDet As New ADODB.Recordset
    Dim rsListaItemsDet2 As New ADODB.Recordset
    Dim i As Integer

    Set rsListaItemsDet = oDCred.CargarOtrasIfis(pnConsCod, pnConsValor)
    Set rsListaItemsDet2 = oDCred.CargarOtrasIfis(pnConsCod, pnConsValor)

    feDetalle.CargaCombo rsListaItemsDet
    If Not (rsListaItemsDet2.EOF And rsListaItemsDet2.BOF) Then
        ReDim vMatListaIfis(rsListaItemsDet2.RecordCount)

        For i = 1 To rsListaItemsDet2.RecordCount
            vMatListaIfis(i - 1) = rsListaItemsDet2!cIfi
            rsListaItemsDet2.MoveNext
        Next
    End If
    Set oDCred = Nothing
End Sub
Private Sub CargarFlexConArrayDetalle()
    Dim Index As Integer, j As Integer
    feDetalle.lbEditarFlex = True
    Call LimpiaFlex(feDetalle)

    For Index = 1 To UBound(fvDetalleRef)
        feDetalle.AdicionaFila
        For j = 0 To UBound(vMatrizListaItem) - 1
            If Trim(Right(vMatrizListaItem(j), 4)) = Trim(Right(fvDetalleRef(Index).cDescripcion, 4)) Then
                feDetalle.TextMatrix(Index, 1) = vMatrizListaItem(j)
                Exit For
            End If
        Next
        feDetalle.TextMatrix(Index, 2) = Format(fvDetalleRef(Index).nImporte, "#,##0.00")
    Next
    SumarMontofeDetalle
End Sub
Private Sub CargarFlexConArrayDetalleDet()
    Dim Index As Integer, j As Integer
    feDetalle.lbEditarFlex = True
    Call LimpiaFlex(feDetalle)

    For Index = 1 To UBound(fvDetalleRef)
        feDetalle.AdicionaFila
        For j = 0 To UBound(vMatListaIfis) - 1
            If Trim(Right(vMatListaIfis(j), 13)) = Trim(Right(fvDetalleRef(Index).cDescripcion, 13)) Then
                feDetalle.TextMatrix(Index, 1) = vMatListaIfis(j)
                Exit For
            End If
        Next
                feDetalle.TextMatrix(Index, 2) = Format(fvDetalleRef(Index).nImporte, "#,##0.00")
    Next
    SumarMontofeDetalle
End Sub
Private Sub feDetalle_Click()
    If Trim(Right(feDetalle.TextMatrix(feDetalle.row, 1), 4)) = 5912 Then 'Que no se edite estado del ejercicio
        Me.feDetalle.ColumnasAEditar = "X-X-X-X-X-X-X"
        Me.feDetalle.ForeColorRow vbBlack, True
        MsgBox "Edición no permitida, valor obtenido de la Utilidad(o Pérdida) Neta", vbInformation, "Aviso"
    Else
        Me.feDetalle.ColumnasAEditar = "X-1-2-X-X-X-X"
        Me.feDetalle.ForeColorRow vbBlack
    End If
    txtTotalDetalle.Text = Format(SumarCampo(feDetalle, 2), "#,##0.00")
End Sub
Private Function CargaDatosMantenimiento(ByVal pnCodEF As Integer, ByVal pnConsCod As Integer, ByVal pnConsValor As Integer, Optional ByVal pnConsValorGrupo As Integer = -1) As Integer
    Dim oCredEval As New COMNCredito.NCOMFormatosEval
    Dim rsListaSubItems As ADODB.Recordset
    Dim nIndice As Integer
    Dim i As Integer
    
    Set rsListaSubItems = oCredEval.RecuperaDatosSubItemDetalle(pnCodEF, pnConsCod, pnConsValor, pnConsValorGrupo)
    If Not (rsListaSubItems.EOF And rsListaSubItems.BOF) Then
        feDetalle.Clear
        FormateaFlex feDetalle
        For nIndice = 1 To rsListaSubItems.RecordCount
            feDetalle.AdicionaFila
            feDetalle.TextMatrix(nIndice, 1) = rsListaSubItems!Concepto
            feDetalle.TextMatrix(nIndice, 2) = Format(rsListaSubItems!nMonto, "#,##0.00")
            rsListaSubItems.MoveNext
        Next
        CargaDatosMantenimiento = nIndice - 1
        
        ReDim Preserve fvDetalleRef(CargaDatosMantenimiento)
        For i = 1 To CargaDatosMantenimiento
            fvDetalleRef(i).cDescripcion = feDetalle.TextMatrix(i, 1)
            fvDetalleRef(i).nImporte = feDetalle.TextMatrix(i, 2)
        Next
    End If
End Function
Private Sub SumarMontofeDetalle()
    Dim i As Integer
    Dim lnMonto As Currency
    Dim lnTotal As Currency
    lnTotal = 0
    If feDetalle.TextMatrix(1, 0) <> "" Then
        For i = 1 To feDetalle.rows - 1
            lnMonto = IIf(IsNumeric(feDetalle.TextMatrix(i, 2)), feDetalle.TextMatrix(i, 2), 0)
            lnTotal = lnTotal + lnMonto
        Next
    End If
    txtTotalDetalle.Enabled = False
    txtTotalDetalle.Text = Format(lnTotal, "#,##0.00")
    fnTotal = Format(lnTotal, "#,##0.00")
End Sub
Private Function ValidaDatos() As String
    Dim nIndice  As Integer
    Dim nMonto As Currency
    ValidaDatos = ""
    
    For nIndice = 1 To feDetalle.rows - 1
        nMonto = IIf(feDetalle.TextMatrix(nIndice, 2) = "", 0, feDetalle.TextMatrix(nIndice, 2))
        If (feDetalle.TextMatrix(nIndice, 1) <> "" Or feDetalle.TextMatrix(nIndice, 2) <> "") And nMonto = 0 Then
            ValidaDatos = "Debe ingresar el monto, en la columna Monto de la fila " & nIndice & " "
            Exit Function
        End If
        If (feDetalle.TextMatrix(nIndice, 1) <> "" Or feDetalle.TextMatrix(nIndice, 2) <> "") Or nMonto > 0 Then
            If feDetalle.TextMatrix(nIndice, 1) = "" Then
                ValidaDatos = "Debe ingresar/seleccionar el detalle del ítem, en la fila " & nIndice & ""
                Exit Function
            End If
        End If
        If ValidaIfiExisteDuplicadoLista(Trim(Right(feDetalle.TextMatrix(nIndice, 1), 4)), nIndice) Then
            ValidaDatos = "No se puede registrar dos veces el mismo detalle del ítem. " & Chr(13) & "Número de fila: " & nIndice & "" & Chr(13) & "Código del ítem:" & Trim(Right(feDetalle.TextMatrix(nIndice, 1), 4)) & " "
            Exit Function
        End If
    Next
End Function
Private Function ValidaIfiExisteDuplicadoLista(ByVal psCodItem As String, ByVal pnFila As Integer) As Boolean
    Dim i As Integer
    ValidaIfiExisteDuplicadoLista = False
    For i = 1 To feDetalle.rows - 1
        If Trim(Right(feDetalle.TextMatrix(i, 1), 4)) = psCodItem Then
            If i <> pnFila Then
                ValidaIfiExisteDuplicadoLista = True
                Exit Function
            End If
        End If
    Next
End Function
Private Sub cmdAgregarDetalle_Click()
    If feDetalle.rows - 1 < 25 Then
        feDetalle.lbEditarFlex = True
        feDetalle.AdicionaFila
        feDetalle.SetFocus
        SendKeys "{Enter}"
    Else
    MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
    End If
End Sub
Private Sub cmdQuitarDetalle_Click()
    If MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        feDetalle.EliminaFila (feDetalle.row)
        txtTotalDetalle.Text = Format(SumarCampo(feDetalle, 2), "#,##0.00")
    End If
End Sub
Private Sub cmdAceptar_Click()
    Dim Index As Integer
    Dim i As Integer
    Dim sMsj As String
    sMsj = ValidaDatos
    
    If sMsj <> "" Then
        MsgBox sMsj, vbInformation, "Alerta"
        Exit Sub
    End If
    
    If MsgBox("¿Desea guardar los valores de los subitems de: " & Chr(13) & "" & fsTitulo & "?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Index = Me.feDetalle.rows - 1
    If feDetalle.TextMatrix(1, 1) = "" Then
        Index = 0
    End If
    
    ReDim Preserve fvDetalleRef(Index)
    If Index > 0 Then
        For i = 1 To Index
            fvDetalleRef(i).cDescripcion = Trim(Right(feDetalle.TextMatrix(i, 1), IIf(fvConsValor = 45, 13, 4)))
            fvDetalleRef(i).nImporte = feDetalle.TextMatrix(i, 2)
        Next i
    End If
    fnTotal = CDbl(txtTotalDetalle.Text)
    Unload Me
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub feDetalle_OnChangeCombo()
    If Trim(Right(feDetalle.TextMatrix(feDetalle.row, 1), 4)) = 451 Then 'Para Ingresar las Ifis
        Me.feDetalle.ColumnasAEditar = "X-1-X-X-X-X-X"
        Me.feDetalle.ForeColorRow vbBlue
        feDetalle.TextMatrix(feDetalle.row, 2) = Format(nTotalSubDetalle, "#,##0.00")
    Else
        Me.feDetalle.ColumnasAEditar = "X-1-2-X-X-X-X"
        Me.feDetalle.ForeColorRow vbBlack
    End If
End Sub
Private Sub feDetalle_OnCellChange(pnRow As Long, pnCol As Long)
    If pnCol = 1 Then
        feDetalle.TextMatrix(feDetalle.row, 1) = UCase(feDetalle.TextMatrix(feDetalle.row, 1))
        If feDetalle.TextMatrix(feDetalle.row, 2) = "" Then
            feDetalle.TextMatrix(feDetalle.row, 2) = "0.00"
        End If
    End If
    If pnCol = 2 Then
        
        If IsNumeric(feDetalle.TextMatrix(pnRow, pnCol)) Then
            If feDetalle.TextMatrix(pnRow, pnCol) < 0 And (fvConsCod = 7041 And fvConsValor <> 19) Then
                feDetalle.TextMatrix(pnRow, pnCol) = "0.00"
            End If
            
            If feDetalle.TextMatrix(pnRow, pnCol) < 0 And (fvConsCod = 7042 And fvConsValor <> 40) Then
                feDetalle.TextMatrix(pnRow, pnCol) = "0.00"
            End If
            
            
            If feDetalle.TextMatrix(pnRow, pnCol) < 0 And (fvConsCod = 7044) Then
                feDetalle.TextMatrix(pnRow, pnCol) = "0.00"
            End If
            
            'If feDetalle.TextMatrix(pnRow, pnCol) < 0 Then
            '    feDetalle.TextMatrix(pnRow, pnCol) = "0.00"
            'End If
        Else
            feDetalle.TextMatrix(pnRow, pnCol) = "0.00"
        End If

        If feDetalle.ColumnasAEditar = "X-1-X-X-X-X-X" Then
            MsgBox "Ingrese el Subdetalle del subItem selecccionado", vbInformation, "Aviso"
        End If
    End If
        txtTotalDetalle.Text = Format(SumarCampo(feDetalle, 2), "#,##0.00")
        
End Sub
Private Sub feDetalle_OnRowChange(pnRow As Long, pnCol As Long)
    feDetalle.TextMatrix(feDetalle.row, 1) = UCase(feDetalle.TextMatrix(feDetalle.row, 1))
End Sub
