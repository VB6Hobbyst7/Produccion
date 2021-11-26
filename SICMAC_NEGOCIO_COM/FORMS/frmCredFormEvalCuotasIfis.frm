VERSION 5.00
Begin VB.Form frmCredFormEvalCuotasIfis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuotas de Ifis"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9105
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCredFormEvalCuotasIfis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fBotones 
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   3250
      Width           =   9135
      Begin VB.CommandButton cmdAceptarIfis 
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
         Left            =   6480
         TabIndex        =   8
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
         Left            =   7800
         TabIndex        =   7
         Top             =   180
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdAgregarIfis 
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
      Left            =   0
      TabIndex        =   3
      Top             =   2880
      Width           =   1170
   End
   Begin VB.CommandButton cmdQuitarIfis 
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
      Left            =   1170
      TabIndex        =   2
      Top             =   2880
      Width           =   1170
   End
   Begin VB.Frame Frame13 
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
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9100
      Begin SICMACT.FlexEdit feCuotaIfis 
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8900
         _ExtentX        =   15690
         _ExtentY        =   4260
         Cols0           =   4
         HighLight       =   1
         EncabezadosNombres=   "N°-Nombre IFI-Monto Cuota-Aux"
         EncabezadosAnchos=   "450-6800-1550-0"
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
         EncabezadosAlineacion=   "C-L-R-C"
         FormatosEdit    =   "0-0-2-2"
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   450
         RowHeight0      =   300
      End
   End
   Begin SICMACT.EditMoney txtTotalIfis 
      Height          =   300
      Left            =   7360
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
      _ExtentX        =   2778
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
      Left            =   6720
      TabIndex        =   5
      Top             =   2895
      Width           =   525
   End
End
Attribute VB_Name = "frmCredFormEvalCuotasIfis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmCredFormEvalIfis
'** Descripción : Formulario que registra las IFIs que se asignan para los formatos de eval respectivos
'** Referencia  : ERS004-2016
'** Creación    : LUCV, 20160528 09:40:00 AM
'**********************************************************************************************
Option Explicit
Dim nTotalCuota As Currency
Dim nTotal As Currency
Dim MatOtraIFIRef As Variant
Dim vMatListaIfis As Variant
Dim lnConsCod As Integer   'CTI320200110 ERS003-2020. Agregó
Dim lnConsValor As Integer 'CTI320200110 ERS003-2020. Agregó
 
'Dim rsListaCompraDeuda As New ADODB.Recordset 'LUCV20161212
'Dim oDCOMInstFinac As New COMDPersona.DCOMInstFinac 'LUCV20161212
'Dim fsCtaCod As String

Private Sub Form_Activate()
  EnfocaControl cmdAgregarIfis
End Sub

Private Sub Form_Load()
    CentraForm Me
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    txtTotalIfis.Enabled = False
    Call cargarControles
End Sub

Public Sub Inicio(Optional ByVal pnMontoFLEX As Long, Optional ByRef pnTotalIFI As Currency = 0, Optional ByRef pMatIFI As Variant, Optional ByVal psTitulo As String, _
                    Optional ByVal pnConsCod As Integer = -1, Optional ByVal pnConsValor As Integer = -1) 'CTI320200110 ERS003-2020. Agregó
    'LUCV20161212, Agregó psCtaCod
    'fsCtaCod = psCtaCod
    
    lnConsCod = pnConsCod     'CTI320200110 ERS003-2020. Agregó
    lnConsValor = pnConsValor 'CTI320200110 ERS003-2020. Agregó
    Call cargarControles 'CTI320200110 ERS003-2020. Agregó
    feCuotaIfis.Clear
    FormateaFlex feCuotaIfis
    If psTitulo <> "" Then
        Me.Caption = psTitulo
    End If
    If IsArray(pMatIFI) Then
        MatOtraIFIRef = pMatIFI
        Call CargarGridConArray
    End If
    If pnMontoFLEX > 0 Then
        MatOtraIFIRef = pMatIFI
        nTotalCuota = pnTotalIFI
    Else
        Set pMatIFI = Nothing
        Set MatOtraIFIRef = Nothing
        nTotalCuota = 0
        pnTotalIFI = 0
        nTotal = 0
    End If
    Me.Show 1
    If IsArray(MatOtraIFIRef) Then
        pMatIFI = MatOtraIFIRef
        pnTotalIFI = nTotal
    End If
End Sub

Private Sub CargarGridConArray()
    Dim i As Integer, j As Integer, nIndice As Integer
    feCuotaIfis.lbEditarFlex = True
    Call LimpiaFlex(feCuotaIfis)
    nTotal = 0
    For i = 0 To UBound(MatOtraIFIRef) - 1
        feCuotaIfis.AdicionaFila
        feCuotaIfis.TextMatrix(i + 1, 0) = MatOtraIFIRef(i, 0)
        
        For j = 0 To UBound(vMatListaIfis) - 1
            feCuotaIfis.TextMatrix(i + 1, 1) = MatOtraIFIRef(i, 1)
            If Trim(Right(vMatListaIfis(j), 13)) = MatOtraIFIRef(i, 1) Then ''LUCV20161115, Modificó->Según ERS068-2016(8-13)
                feCuotaIfis.TextMatrix(i + 1, 1) = vMatListaIfis(j)
                Exit For
            End If
        Next
        feCuotaIfis.TextMatrix(i + 1, 2) = MatOtraIFIRef(i, 2)
        nTotal = nTotal + feCuotaIfis.TextMatrix(i + 1, 2)
    Next i
    
    nIndice = IIf(feCuotaIfis.TextMatrix(1, 1) = "", 0, feCuotaIfis.rows - 1)
    ReDim MatOtraIFIRef(nIndice, 3)
    If nIndice > 0 Then
        For i = 1 To feCuotaIfis.rows - 1
            MatOtraIFIRef(i - 1, 0) = feCuotaIfis.TextMatrix(i, 0)
            MatOtraIFIRef(i - 1, 1) = Trim(Right(feCuotaIfis.TextMatrix(i, 1), 13)) 'LUCV20161115, Modificó->Según ERS068-2016
            MatOtraIFIRef(i - 1, 2) = feCuotaIfis.TextMatrix(i, 2)
        Next i
    End If
    txtTotalIfis.Text = Format(nTotal, "#,##0.00")
End Sub
Private Sub cmdAceptarIfis_Click()
    Dim sMsj As String
    
    sMsj = ValidaDatosGrabar
    
    If sMsj <> "" Then
        MsgBox sMsj, vbInformation, "Alerta"
        Exit Sub
    End If
        
    If MsgBox("Desea Guardar IFIs??", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        
        nTotal = CCur(txtTotalIfis.Text)
        Dim i As Integer ' Llenado de Matriz
        Dim nIndice As Integer
        
        'Call ValidaDatosGrabar
        nIndice = IIf(feCuotaIfis.TextMatrix(1, 1) = "", 0, feCuotaIfis.rows - 1)
        ReDim MatOtraIFIRef(nIndice, 3)
        If nIndice > 0 Then
            For i = 1 To feCuotaIfis.rows - 1
                MatOtraIFIRef(i - 1, 0) = feCuotaIfis.TextMatrix(i, 0)
                MatOtraIFIRef(i - 1, 1) = Trim(Right(feCuotaIfis.TextMatrix(i, 1), 13)) 'LUCV20161115, Modificó->Según ERS068-2016
                MatOtraIFIRef(i - 1, 2) = feCuotaIfis.TextMatrix(i, 2)
            Next i
        End If
    Unload Me
End Sub

Private Sub cmdAgregarIfis_Click()
    If feCuotaIfis.rows - 1 < 25 Then
        feCuotaIfis.lbEditarFlex = True
        feCuotaIfis.AdicionaFila
        feCuotaIfis.SetFocus
        SendKeys "{Enter}"
    Else
    MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdQuitarIfis_Click()
    If MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        feCuotaIfis.EliminaFila (feCuotaIfis.row)
        txtTotalIfis.Text = Format(SumarCampo(feCuotaIfis, 2), "#,##0.00")
    End If
End Sub

Private Sub feCuotaIfis_OnCellChange(pnRow As Long, pnCol As Long)
    Select Case pnCol
    Case 1
        feCuotaIfis.TextMatrix(pnRow, pnCol) = UCase(feCuotaIfis.TextMatrix(pnRow, pnCol))
        If ValidaIfiExisteDuplicadoLista(Trim(Right(feCuotaIfis.TextMatrix(pnRow, pnCol), 13)), pnRow) Then 'LUCV20161115, Modificó->Según ERS068-2016
            MsgBox "No se puede registrar dos veces una misma IFI", vbInformation, "Alerta"
            feCuotaIfis.TextMatrix(pnRow, 1) = ""
            feCuotaIfis.TextMatrix(pnRow, 2) = ""
        End If
    Case 2
        If IsNumeric(feCuotaIfis.TextMatrix(pnRow, pnCol)) Then
            If feCuotaIfis.TextMatrix(pnRow, pnCol) < 0 Then
                feCuotaIfis.TextMatrix(pnRow, pnCol) = "0.00"
            End If
        Else
            feCuotaIfis.TextMatrix(pnRow, pnCol) = "0.00"
        End If
    End Select
    txtTotalIfis.Text = Format(SumarCampo(feCuotaIfis, 2), "#,##0.00")
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cargarControles()
    Dim oDCred As New COMDCredito.DCOMFormatosEval
    Dim rsLista As New ADODB.Recordset
    Dim rsLista2 As New ADODB.Recordset
    Dim i As Integer
    
    Set rsLista = oDCred.CargarOtrasIfis(lnConsCod, lnConsValor)
    Set rsLista2 = oDCred.CargarOtrasIfis(lnConsCod, lnConsValor)
   
    feCuotaIfis.CargaCombo rsLista
    If Not (rsLista2.EOF And rsLista2.BOF) Then
        ReDim vMatListaIfis(rsLista2.RecordCount)
        
        For i = 1 To rsLista2.RecordCount
            vMatListaIfis(i - 1) = rsLista2!cIfi
            rsLista2.MoveNext
        Next
    End If
    Set oDCred = Nothing
End Sub

Private Function ValidaIfiExisteDuplicadoLista(ByVal psCodIfi As String, ByVal pnFila As Integer) As Boolean
    Dim i As Integer
    
    ValidaIfiExisteDuplicadoLista = False
    
    For i = 1 To feCuotaIfis.rows - 1
        If Trim(Right(feCuotaIfis.TextMatrix(i, 1), 13)) = psCodIfi Then 'LUCV20161115, Modificó->Según ERS068-2016
            If i <> pnFila Then
                ValidaIfiExisteDuplicadoLista = True
                Exit Function
            End If
        End If
    Next
End Function

Private Function ValidaDatosGrabar() As String
    Dim i As Integer

    For i = 1 To feCuotaIfis.rows - 1
        
        If feCuotaIfis.TextMatrix(i, 1) <> "" Then
            If val(Replace(feCuotaIfis.TextMatrix(i, 2), ",", "")) = 0 Then
                ValidaDatosGrabar = "Debe ingresar un monto"
                Exit Function
            End If
        End If
        
        If feCuotaIfis.TextMatrix(i, 0) <> "" Then
            If feCuotaIfis.TextMatrix(i, 1) = "" Then
                ValidaDatosGrabar = "No se pueden agregar datos vacíos. Coordinar con el Dpto. de TI."
                Exit Function
            End If
        End If
        
        If ValidaIfiExisteDuplicadoLista(Trim(Right(feCuotaIfis.TextMatrix(i, 1), 13)), i) Then 'LUCV20161115, Modificó->Según ERS068-2016
            ValidaDatosGrabar = "No se puede registrar dos veces una misma IFI"
            Exit Function
        End If
        
        If feCuotaIfis.TextMatrix(i, 0) <> "" Then
            If Left(feCuotaIfis.TextMatrix(i, 1), 7) = "SIN IFI" Then
                ValidaDatosGrabar = "El código de IFI no tiene correspondencia con el negocio. Coordinar con el Dpto. de TI."
                Exit Function
            End If
        End If
       
    Next

End Function


