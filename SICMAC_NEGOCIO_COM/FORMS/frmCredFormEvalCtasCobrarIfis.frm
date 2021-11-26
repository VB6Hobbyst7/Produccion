VERSION 5.00
Begin VB.Form frmCredFormEvalCtasCobrarIfis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Titulo"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   Icon            =   "frmCredFormEvalCtasCobrarIfis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   9330
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
      Left            =   1400
      TabIndex        =   1
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
      TabIndex        =   6
      Top             =   0
      Width           =   9250
      Begin SICMACT.FlexEdit feCtasCobrar 
         Height          =   2175
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   9015
         _extentx        =   15901
         _extenty        =   3836
         cols0           =   5
         highlight       =   1
         encabezadosnombres=   "-Fecha-Deudas IFI'S-Total-aux"
         encabezadosanchos=   "0-950-6780-1200-0"
         font            =   "frmCredFormEvalCtasCobrarIfis.frx":030A
         font            =   "frmCredFormEvalCtasCobrarIfis.frx":0332
         font            =   "frmCredFormEvalCtasCobrarIfis.frx":035A
         font            =   "frmCredFormEvalCtasCobrarIfis.frx":0382
         font            =   "frmCredFormEvalCtasCobrarIfis.frx":03AA
         fontfixed       =   "frmCredFormEvalCtasCobrarIfis.frx":03D2
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1  'True
         tipobusqueda    =   3
         columnasaeditar =   "X-1-2-3-X"
         listacontroles  =   "0-2-3-0-0"
         encabezadosalineacion=   "L-L-L-R-C"
         formatosedit    =   "0-0-0-2-2"
         lbeditarflex    =   -1  'True
         lbbuscaduplicadotext=   -1  'True
         rowheight0      =   300
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   3000
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
         Left            =   6750
         TabIndex        =   2
         Top             =   180
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
         Top             =   180
         Width           =   1170
      End
   End
   Begin SICMACT.EditMoney txtTotalCtasCobrar 
      Height          =   300
      Left            =   7560
      TabIndex        =   7
      Top             =   2700
      Width           =   1500
      _extentx        =   2646
      _extenty        =   529
      font            =   "frmCredFormEvalCtasCobrarIfis.frx":03F8
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
      Left            =   6720
      TabIndex        =   8
      Top             =   2700
      Width           =   525
   End
End
Attribute VB_Name = "frmCredFormEvalCtasCobrarIfis"
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
Dim vMatListaIfis As Variant
Dim fsTitulo As String

Private Sub Form_Load()
    CentraForm Me
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    'Call CargarControles(fvConsCod, fvConsValor)'Comentó->Según ERS068-2016
End Sub

Public Function Inicio(ByRef pvDetalleActivoFlex() As tFormEvalDetalleActivosCtasCobrarFormato5, ByRef pnTotalCeldaFlex As Double, _
                       ByRef pnConsCod As Integer, ByRef pnConsValor As Integer, ByVal psTitulo As String, Optional ByVal psCtaCod As String, _
                       Optional ByVal pnTpoPat As Integer) As Boolean
    Me.Caption = psTitulo
    fsTitulo = psTitulo
    fvConsCod = pnConsCod 'LUCV20161115, Agregó->Según ERS068-2016
    fvConsValor = pnConsValor 'LUCV20161115, Agregó->Según ERS068-2016
    Call CargarControles(fvConsCod, fvConsValor) 'LUCV20161115, Agregó->Según ERS068-2016
    
    If IsArray(pvDetalleActivoFlex) Then
        If UBound(pvDetalleActivoFlex) > 0 Then 'Si Matrix Contiene Datos
            fvDetalleRef = pvDetalleActivoFlex
            Call SetFlexDetalleCtaCobrar
            fnTotal = pnTotalCeldaFlex
            'fvConsCod = pnConsCod 'LUCV20161115, Comentó->Según ERS068-2016
            'fvConsValor = pnConsValor 'LUCV20161115, Comentó->Según ERS068-2016
            Call SumarMontofeCtasCobrar
        Else
            ReDim pvDetalleActivoFlex(CargaDatosMantenimiento(psCtaCod, pnConsCod, pnConsValor, pnTpoPat))
            'fvConsCod = pnConsCod 'LUCV20161115, Comentó->Según ERS068-2016
            'fvConsValor = pnConsValor 'LUCV20161115, Comentó->Según ERS068-2016
            Call SumarMontofeCtasCobrar
        End If
        
    Else
        ReDim pvDetalleActivoFlex(0)
        pnTotalCeldaFlex = 0
        fnTotal = 0
    End If
        
    Me.Show 1
    pvDetalleActivoFlex = fvDetalleRef
    pnTotalCeldaFlex = fnTotal
    pnConsCod = fvConsValor
    pnConsValor = fvConsValor
    Inicio = True
    
End Function

Private Sub SetFlexDetalleCtaCobrar()
    Dim index As Integer, j As Integer
    feCtasCobrar.lbEditarFlex = True
    Call LimpiaFlex(feCtasCobrar)

    For index = 1 To UBound(fvDetalleRef)
            feCtasCobrar.AdicionaFila
            feCtasCobrar.TextMatrix(index, 1) = fvDetalleRef(index).dFecha
            For j = 0 To UBound(vMatListaIfis) - 1
                If Trim(Right(vMatListaIfis(j), 13)) = fvDetalleRef(index).cCtaporCobrar Then 'LUCV20161115, Modificó->Según ERS068-2016(8-13)
                    feCtasCobrar.TextMatrix(index, 2) = vMatListaIfis(j)
                    Exit For
                End If
            Next
            feCtasCobrar.TextMatrix(index, 3) = Format(fvDetalleRef(index).nTotal, "#,##0.00")
    Next
    SumarMontofeCtasCobrar
End Sub
Private Sub SumarMontofeCtasCobrar()
    Dim I As Integer
    Dim lnMonto As Currency
    Dim lnTotal As Currency
    lnTotal = 0
    If feCtasCobrar.TextMatrix(1, 0) <> "" Then
        For I = 1 To feCtasCobrar.Rows - 1
            lnMonto = IIf(IsNumeric(feCtasCobrar.TextMatrix(I, 3)), feCtasCobrar.TextMatrix(I, 3), 0)
            lnTotal = lnTotal + lnMonto
        Next
    End If
    txtTotalCtasCobrar.Enabled = False
    txtTotalCtasCobrar.Text = Format(lnTotal, "#,##0.00")
    fnTotal = Format(lnTotal, "#,##0.00")
End Sub

Private Sub cmdAceptar_Click()
    Dim index As Integer
    Dim I As Integer
    Dim sMsj As String
    sMsj = ValidaDatos
    
    If sMsj <> "" Then
        MsgBox sMsj, vbInformation, "Alerta"
        Exit Sub
    End If
    
    If MsgBox("¿Desea guardar, " & fsTitulo & "?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    index = Me.feCtasCobrar.Rows - 1
    If feCtasCobrar.TextMatrix(1, 1) = "" Then
        index = 0
    End If
    
    ReDim Preserve fvDetalleRef(index)
    If index > 0 Then
        For I = 1 To index
            fvDetalleRef(I).dFecha = feCtasCobrar.TextMatrix(I, 1)
            fvDetalleRef(I).cCtaporCobrar = Trim(Right(feCtasCobrar.TextMatrix(I, 2), 13)) 'LUCV20161115, Modificó->Según ERS068-2016(8-13)
            fvDetalleRef(I).nTotal = feCtasCobrar.TextMatrix(I, 3)
        Next I
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
    If pnCol = 2 Then
    feCtasCobrar.TextMatrix(feCtasCobrar.row, 2) = UCase(feCtasCobrar.TextMatrix(feCtasCobrar.row, 2))
    End If
    If pnCol = 3 Then
    If IsNumeric(feCtasCobrar.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos

        If feCtasCobrar.TextMatrix(pnRow, pnCol) < 0 Then
            feCtasCobrar.TextMatrix(pnRow, pnCol) = "0.00"
        End If
    Else
        feCtasCobrar.TextMatrix(pnRow, pnCol) = "0.00"
    End If
    End If
    txtTotalCtasCobrar.Text = Format(SumarCampo(feCtasCobrar, 3), "#,##0.00")
End Sub

Private Sub feCtasCobrar_OnRowChange(pnRow As Long, pnCol As Long)
    feCtasCobrar.TextMatrix(feCtasCobrar.row, 1) = UCase(feCtasCobrar.TextMatrix(feCtasCobrar.row, 1))
End Sub

Private Sub CargarControles(Optional pnConsCod As Integer = -1, Optional pnConsValor As Integer = -1) 'LUCV20161115, Modificó->Según ERS068-2016(Agregó parametros)
   
    Dim oDCred As New COMDCredito.DCOMFormatosEval
    Dim rsLista As New ADODB.Recordset
    Dim rsLista2 As New ADODB.Recordset
    Dim I As Integer
    
    Set rsLista = oDCred.CargarOtrasIfis(pnConsCod, pnConsValor)
    Set rsLista2 = oDCred.CargarOtrasIfis(pnConsCod, pnConsValor)
    
    feCtasCobrar.CargaCombo rsLista
    
    If Not (rsLista2.EOF And rsLista2.BOF) Then
        ReDim vMatListaIfis(rsLista2.RecordCount)
        
        For I = 1 To rsLista2.RecordCount
            vMatListaIfis(I - 1) = rsLista2!cIfi
            rsLista2.MoveNext
        Next
    End If
    Set oDCred = Nothing
End Sub

Private Function ValidaDatos() As String
    Dim nIndice  As Integer
    Dim nMonto As Currency
    ValidaDatos = ""
    
    For nIndice = 1 To feCtasCobrar.Rows - 1
        nMonto = IIf(feCtasCobrar.TextMatrix(nIndice, 3) = "", 0, feCtasCobrar.TextMatrix(nIndice, 3))
        If (feCtasCobrar.TextMatrix(nIndice, 1) <> "" Or feCtasCobrar.TextMatrix(nIndice, 2) <> "") And nMonto = 0 Then
            ValidaDatos = "Debe ingresar un monto"
            Exit Function
        End If
        If (feCtasCobrar.TextMatrix(nIndice, 1) <> "" Or feCtasCobrar.TextMatrix(nIndice, 2) <> "") Or nMonto > 0 Then
            If feCtasCobrar.TextMatrix(nIndice, 1) = "" Then
                ValidaDatos = "Debe ingresar una fecha"
                Exit Function
            End If
            If feCtasCobrar.TextMatrix(nIndice, 2) = "" Then
                ValidaDatos = "Debe ingresar una Ifi"
                Exit Function
            End If
        End If
        
        If ValidaIfiExisteDuplicadoLista(Trim(Right(feCtasCobrar.TextMatrix(nIndice, 2), 13)), nIndice) Then 'LUCV20161115, Modificó->Según ERS068-2016(8-13)
            ValidaDatos = "No se puede registrar dos veces una misma IFI"
            Exit Function
        End If
    Next
End Function

Private Function CargaDatosMantenimiento(ByVal psCtaCod As String, ByVal pnConsCod As Integer, ByVal pnConsValor As Integer, ByVal pnTpoPat As Integer) As Integer
    Dim oCredEval As New COMNCredito.NCOMFormatosEval
    Dim oRS As ADODB.Recordset
    Dim nIndice As Integer
    Dim I As Integer
    
    Set oRS = oCredEval.RecuperaDatosCtaCobrar(psCtaCod, pnConsCod, pnConsValor, pnTpoPat)
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
        For I = 1 To CargaDatosMantenimiento
            fvDetalleRef(I).dFecha = feCtasCobrar.TextMatrix(I, 1)
            fvDetalleRef(I).cCtaporCobrar = Trim(Right(feCtasCobrar.TextMatrix(I, 2), 13)) 'LUCV20161115, Modificó->Según ERS068-2016(8-13)
            fvDetalleRef(I).nTotal = feCtasCobrar.TextMatrix(I, 3)
        Next
    End If
    
End Function

Private Function ValidaIfiExisteDuplicadoLista(ByVal psCodIfi As String, ByVal pnFila As Integer) As Boolean
    Dim I As Integer
    
    ValidaIfiExisteDuplicadoLista = False
    
    For I = 1 To feCtasCobrar.Rows - 1
        If Trim(Right(feCtasCobrar.TextMatrix(I, 2), 13)) = psCodIfi Then 'LUCV20161115, Modificó->Según ERS068-2016(8-13)
            If I <> pnFila Then
                ValidaIfiExisteDuplicadoLista = True
                Exit Function
            End If
        End If
    Next
End Function


