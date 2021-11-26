VERSION 5.00
Begin VB.Form frmCredFormEvalCtasCobrar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ctas por Cobras"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCredFormEvalCtasCobrar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   3000
      Width           =   7335
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
         Left            =   6000
         TabIndex        =   3
         Top             =   160
         Width           =   1170
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
         Left            =   4800
         TabIndex        =   2
         Top             =   160
         Width           =   1170
      End
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
      TabIndex        =   4
      Top             =   0
      Width           =   7425
      Begin SICMACT.FlexEdit feCtasCobrar 
         Height          =   2175
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   3836
         Cols0           =   5
         HighLight       =   1
         EncabezadosNombres=   "N-Fecha-Cuentas por Cobrar-Total-aux"
         EncabezadosAnchos=   "350-1060-4000-1700-0"
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
         ColumnasAEditar =   "X-1-2-3-X"
         ListaControles  =   "0-2-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-L-L-R-C"
         FormatosEdit    =   "0-0-0-2-2"
         TextArray0      =   "N"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
      End
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
      Top             =   2640
      Width           =   1170
   End
   Begin SICMACT.EditMoney txtTotalCtasCobrar 
      Height          =   300
      Left            =   5760
      TabIndex        =   7
      Top             =   2700
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
      Left            =   5040
      TabIndex        =   8
      Top             =   2700
      Width           =   525
   End
End
Attribute VB_Name = "frmCredFormEvalCtasCobrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmCredFormEvalCtasCobrar
'** Descripción : Formulario para evaluación de Creditos que tienen el tipo de evaluación 2
'** Referencia  : ERS004-2016
'** Creación    : LUCV, 20160525 09:00:00 AM
'**********************************************************************************************
Option Explicit
Dim fvDetalleRef() As tFormEvalDetalleActivosCtasCobrarFormato5
Dim fnNroFila As Integer
Dim fnTotal As Double
Dim fvConsCod As Integer
Dim fvConsValor As Integer
Dim fnTpoPat As Integer
Dim fsTitulo As String

Private Sub Form_Load()
CentraForm Me
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
End Sub

Public Function Inicio(ByRef pvDetalleActivoFlex() As tFormEvalDetalleActivosCtasCobrarFormato5, ByRef pnTotalCeldaFlex As Double, ByRef pnConsCod As Integer, _
                       ByRef pnConsValor As Integer, ByVal psTitulo As String, Optional ByVal psCtaCod As String, Optional ByVal pnTpoPat As Integer) As Boolean
    Me.Caption = psTitulo
    fsTitulo = psTitulo
    fnTpoPat = pnTpoPat
    If UBound(pvDetalleActivoFlex) > 0 Then 'Si Matrix Contiene Datos
        fvDetalleRef = pvDetalleActivoFlex
        Call SetFlexDetalleCtaCobrar
        fnTotal = pnTotalCeldaFlex
        fvConsCod = pnConsCod
        fvConsValor = pnConsValor
        Call SumarMontofeCtasCobrar
    Else
        ReDim pvDetalleActivoFlex(CargaDatosMantenimiento(psCtaCod, pnConsCod, pnConsValor, pnTpoPat))
        fvConsCod = pnConsCod
        fvConsValor = pnConsValor
        Call SumarMontofeCtasCobrar
    End If
        
    Me.Show 1
    pvDetalleActivoFlex = fvDetalleRef
    pnTotalCeldaFlex = fnTotal
    pnConsCod = fvConsValor
    pnConsValor = fvConsValor
    Inicio = True
    
End Function

Private Sub SetFlexDetalleCtaCobrar()
    Dim index As Integer
    feCtasCobrar.lbEditarFlex = True
    Call LimpiaFlex(feCtasCobrar)

    For index = 1 To UBound(fvDetalleRef)
            feCtasCobrar.AdicionaFila
            feCtasCobrar.TextMatrix(index, 1) = fvDetalleRef(index).dFecha
            feCtasCobrar.TextMatrix(index, 2) = fvDetalleRef(index).cCtaporCobrar
            feCtasCobrar.TextMatrix(index, 3) = Format(fvDetalleRef(index).nTotal, "#,##0.00")
    Next
    SumarMontofeCtasCobrar
End Sub
Private Sub SumarMontofeCtasCobrar()
    Dim i As Integer
    Dim lnMonto As Currency
    Dim lnTotal As Currency
    lnTotal = 0
    If feCtasCobrar.TextMatrix(1, 0) <> "" Then
        For i = 1 To feCtasCobrar.Rows - 1
            lnMonto = IIf(IsNumeric(feCtasCobrar.TextMatrix(i, 3)), feCtasCobrar.TextMatrix(i, 3), 0)
            lnTotal = lnTotal + lnMonto
        Next
    End If
    txtTotalCtasCobrar.Enabled = False
    txtTotalCtasCobrar.Text = Format(lnTotal, "#,##0.00")
    fnTotal = Format(lnTotal, "#,##0.00")
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
    If MsgBox("¿Desea Guardar, " & fsTitulo & "?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub

        index = IIf(feCtasCobrar.TextMatrix(1, 3) = "", 0, Me.feCtasCobrar.Rows - 1)

        ReDim Preserve fvDetalleRef(index)
        If index > 0 Then
            For i = 1 To index
                fvDetalleRef(i).dFecha = Format(feCtasCobrar.TextMatrix(i, 1), "")
                fvDetalleRef(i).cCtaporCobrar = feCtasCobrar.TextMatrix(i, 2)
                fvDetalleRef(i).nTotal = feCtasCobrar.TextMatrix(i, 3)
            Next i
        End If
        fnTotal = CDbl(txtTotalCtasCobrar.Text)
    Unload Me
End Sub

Private Sub cmdAgregar_Click()
    If feCtasCobrar.Rows - 1 < 25 Then
        feCtasCobrar.lbEditarFlex = True
        feCtasCobrar.AdicionaFila
        feCtasCobrar.SetFocus
        SendKeys "{Enter}"
    Else
    MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdQuitar_Click()
    If MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        feCtasCobrar.EliminaFila (feCtasCobrar.row)
        txtTotalCtasCobrar.Text = Format(SumarCampo(feCtasCobrar, 3), "#,##0.00")
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub feCtasCobrar_OnCellChange(pnRow As Long, pnCol As Long)
    Select Case pnCol
    Case 2
        feCtasCobrar.TextMatrix(feCtasCobrar.row, 2) = UCase(feCtasCobrar.TextMatrix(feCtasCobrar.row, 2))
    Case 3
        If IsNumeric(feCtasCobrar.TextMatrix(pnRow, pnCol)) Then
            If feCtasCobrar.TextMatrix(pnRow, pnCol) >= 0 Then
                txtTotalCtasCobrar.Text = Format(SumarCampo(feCtasCobrar, 3), "#,##0.00")
            Else
                feCtasCobrar.TextMatrix(pnRow, pnCol) = "0.00"
            End If
        Else
            feCtasCobrar.TextMatrix(pnRow, pnCol) = "0.00"
            txtTotalCtasCobrar.Text = Format(SumarCampo(feCtasCobrar, 3), "#,##0.00")
        End If
    End Select
End Sub

Private Sub feCtasCobrar_OnRowChange(pnRow As Long, pnCol As Long)
    feCtasCobrar.TextMatrix(feCtasCobrar.row, 1) = UCase(feCtasCobrar.TextMatrix(feCtasCobrar.row, 1))
End Sub

Private Function ValidaCamposVacios() As String
    Dim nIndice As Integer
    ValidaCamposVacios = ""
    If feCtasCobrar.TextMatrix(0, 1) <> "" Then
        For nIndice = 1 To feCtasCobrar.Rows - 1
            If feCtasCobrar.TextMatrix(nIndice, 1) = "" Or feCtasCobrar.TextMatrix(nIndice, 2) = "" Or feCtasCobrar.TextMatrix(nIndice, 3) = "" Then
                ValidaCamposVacios = "No se puede registrar valores vacios"
            Else
                If val(Replace(feCtasCobrar.TextMatrix(nIndice, 3), ",", "")) = 0 Then
                    ValidaCamposVacios = "Debe ingresar valores mayores a 0"
                End If
            End If
        Next
    End If
End Function

Private Function CargaDatosMantenimiento(ByVal psCtaCod As String, ByVal pnConsCod As Integer, ByVal pnConsValor As Integer, ByVal pnTpoPat As Integer) As Integer
    Dim oCredEval As New COMNCredito.NCOMFormatosEval
    Dim oRS As ADODB.Recordset
    Dim nIndice As Integer
    Dim i As Integer
    
    Set oRS = oCredEval.RecuperaDatosCtaCobrar(psCtaCod, pnConsCod, pnConsValor, pnTpoPat, 2)
    If Not (oRS.EOF And oRS.BOF) Then
        feCtasCobrar.Clear
        FormateaFlex feCtasCobrar
        For nIndice = 1 To oRS.RecordCount
            feCtasCobrar.AdicionaFila
            feCtasCobrar.TextMatrix(nIndice, 1) = Format(oRS!dCtaFecha, "DD/MM/YYYY")
            feCtasCobrar.TextMatrix(nIndice, 2) = oRS!cDescripcion
            feCtasCobrar.TextMatrix(nIndice, 3) = Format(oRS!nTotal, "#,##0.00")
            oRS.MoveNext
        Next
        CargaDatosMantenimiento = nIndice - 1
        
        ReDim Preserve fvDetalleRef(CargaDatosMantenimiento)
        For i = 1 To CargaDatosMantenimiento
            fvDetalleRef(i).dFecha = feCtasCobrar.TextMatrix(i, 1)
            'fvDetalleRef(i).cCtaporCobrar = Trim(Right(feCtasCobrar.TextMatrix(i, 2), 13)) 'LUCV20161115, Comentó y modificó->Según ERS068-2016
            fvDetalleRef(i).cCtaporCobrar = Trim(feCtasCobrar.TextMatrix(i, 2)) 'LUCV20161115, Comentó y modificó->Según ERS068-2016
            fvDetalleRef(i).nTotal = feCtasCobrar.TextMatrix(i, 3)
        Next
    End If
    
End Function

