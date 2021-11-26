VERSION 5.00
Begin VB.Form frmInvActivarBienComponente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Componentes"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9480
   Icon            =   "frmInvActivarBienComponente.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   2320
      Width           =   1695
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7320
      TabIndex        =   4
      Top             =   2315
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8380
      TabIndex        =   3
      Top             =   2315
      Width           =   1050
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   80
      TabIndex        =   2
      Top             =   2315
      Width           =   1050
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "&Quitar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1150
      TabIndex        =   1
      Top             =   2315
      Width           =   1050
   End
   Begin Sicmact.FlexEdit feComponente 
      Height          =   2205
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   3889
      Cols0           =   12
      HighLight       =   1
      AllowUserResizing=   1
      EncabezadosNombres=   "#-Tipo-Objeto-Nombre-Cod. Inventario-Depc.Cont (meses)-Depc.Trib (meses)-Marca-Modelo-Serie-Precio-Aux"
      EncabezadosAnchos=   "350-1800-1500-2500-1500-1750-1750-1500-1500-1500-1200-0"
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
      ColumnasAEditar =   "X-1-2-X-X-5-X-7-8-9-10-X"
      ListaControles  =   "0-3-1-0-0-0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-L-C-L-L-R-R-L-L-L-R-C"
      FormatosEdit    =   "0-0-0-0-0-3-3-0-0-0-2-0"
      CantEntero      =   9
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbPuntero       =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   345
      RowHeight0      =   300
   End
   Begin VB.Label Label1 
      Caption         =   "Total:"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   2355
      Width           =   495
   End
End
Attribute VB_Name = "frmInvActivarBienComponente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************
'** Nombre : frmInvActivarBienComponente
'** Descripción : Registro de Componentes Activo Fijo creado segun ERS059-2013
'** Creación : EJVG, 20130603 09:00:00 AM
'*****************************************************************************
Option Explicit
Dim fbAceptar As Boolean
Dim fsCodInventarioUltimoAF As String
Dim fsCodInventarioUltimoBND As String
Dim fsAgeCod As String

Private Sub Form_Load()
    CentraForm Me
End Sub
Public Function Inicio(ByVal pMatComponente As Variant, ByVal psCodInventarioAF As String, ByVal psCodInventarioBND As String, ByVal psAgeCod As String) As Variant
    Dim Mat As Variant
    Dim i As Integer
    Dim lnTotal As Currency
    fsCodInventarioUltimoAF = psCodInventarioAF
    fsCodInventarioUltimoBND = psCodInventarioBND
    fsAgeCod = psAgeCod
    'Llena Detalle ***
    Call LimpiaFlex(feComponente)
    For i = 1 To UBound(pMatComponente, 2)
        feComponente.AdicionaFila
        feComponente.TextMatrix(feComponente.Row, 1) = IIf(pMatComponente(1, i) = "1", "ACTIVO FIJO" & Space(75) & "1", "BIEN NO DEPREC." & Space(75) & "2") 'Tipo
        feComponente.TextMatrix(feComponente.Row, 2) = pMatComponente(2, i) 'Objeto
        feComponente.TextMatrix(feComponente.Row, 3) = pMatComponente(3, i) 'Nombre Objeto
        feComponente.TextMatrix(feComponente.Row, 4) = pMatComponente(4, i) 'Cod. Inventario
        feComponente.TextMatrix(feComponente.Row, 5) = pMatComponente(5, i) 'Depreciacion Contable (meses)
        feComponente.TextMatrix(feComponente.Row, 6) = pMatComponente(6, i) 'Depreciacion Tributaria (meses)
        feComponente.TextMatrix(feComponente.Row, 7) = pMatComponente(7, i) 'Marca
        feComponente.TextMatrix(feComponente.Row, 8) = pMatComponente(8, i) 'Modelo
        feComponente.TextMatrix(feComponente.Row, 9) = pMatComponente(9, i) 'Serie
        feComponente.TextMatrix(feComponente.Row, 10) = pMatComponente(10, i) 'Precio
        lnTotal = lnTotal + CCur(pMatComponente(10, i))
    Next
    txtTotal.Text = Format(lnTotal, gsFormatoNumeroView)
    feComponente.TopRow = 1
    Show 1
    ReDim Mat(1 To 10, 0)
    'Recupera Detalle ***
    If fbAceptar = True Then
        If Not FlexVacio(feComponente) Then
            For i = 1 To feComponente.Rows - 1
                ReDim Preserve Mat(1 To 10, 1 To i)
                Mat(1, i) = CInt(Trim(Right(feComponente.TextMatrix(i, 1), 2))) 'Tipo
                Mat(2, i) = Trim(feComponente.TextMatrix(i, 2)) 'Objeto
                Mat(3, i) = Trim(feComponente.TextMatrix(i, 3)) 'Nombre Objeto
                Mat(4, i) = Trim(feComponente.TextMatrix(i, 4)) 'Cod. Inventario
                Mat(5, i) = Trim(feComponente.TextMatrix(i, 5)) 'Depreciacion Contable (meses)
                Mat(6, i) = Trim(feComponente.TextMatrix(i, 6)) 'Depreciacion Tributaria (meses)
                Mat(7, i) = Trim(feComponente.TextMatrix(i, 7)) 'Marca
                Mat(8, i) = Trim(feComponente.TextMatrix(i, 8)) 'Modelo
                Mat(9, i) = Trim(feComponente.TextMatrix(i, 9)) 'Serie
                Mat(10, i) = CCur(Trim(feComponente.TextMatrix(i, 10))) 'Precio
            Next
        End If
        Inicio = Mat
    Else
        Inicio = pMatComponente
    End If
End Function
Private Sub cmdAgregar_Click()
    If Not ValidarIngresoDatosFlex Then Exit Sub
    feComponente.AdicionaFila
    feComponente.SetFocus
    feComponente.Col = 1
    feComponente_RowColChange
End Sub
Private Sub cmdQuitar_Click()
    feComponente.EliminaFila feComponente.Row
End Sub
Private Sub cmdAceptar_Click()
    Dim i As Long
    If Not ValidarIngresoDatosFlex Then Exit Sub
    If Not FlexVacio(feComponente) Then
        For i = 1 To feComponente.Rows - 1
            If (CInt(Trim(Right(feComponente.TextMatrix(i, 1), 2))) = 1) Then
                If Val(feComponente.TextMatrix(i, 5)) = 0 Then
                    MsgBox "No se puede continuar porque el Objeto " & UCase(feComponente.TextMatrix(i, 3)) & " no tiene especificado la Depreciación Contable", vbCritical, "Aviso"
                    feComponente.TopRow = i
                    feComponente.Row = i
                    Exit Sub
                End If
                If Val(feComponente.TextMatrix(i, 6)) = 0 Then
                    MsgBox "No se puede continuar porque el Objeto " & UCase(feComponente.TextMatrix(i, 3)) & " no tiene especificada la Depreciación Tributaria", vbCritical, "Aviso"
                    feComponente.TopRow = i
                    feComponente.Row = i
                    Exit Sub
                End If
            End If
        Next
    End If
    fbAceptar = True
    Hide
End Sub
Private Sub cmdCancelar_Click()
    fbAceptar = False
    Hide
End Sub
Private Sub feComponente_RowColChange()
    Dim rsOpt As New ADODB.Recordset
    If feComponente.Col = 1 Then
        With rsOpt
            .Fields.Append "desc", adVarChar, 50
            .Fields.Append "value", adVarChar, 2
            .Open
            .AddNew
            .Fields("desc") = "ACTIVO FIJO"
            .Fields("value") = "1"
            .AddNew
            .Fields("desc") = "BIEN NO DEPREC."
            .Fields("value") = "2"
        End With
        rsOpt.MoveFirst
        feComponente.CargaCombo rsOpt
    End If
    Set rsOpt = Nothing
End Sub
Private Sub feComponente_OnCellChange(pnRow As Long, pnCol As Long)
    Dim i As Integer
    Dim lnTotal As Currency
    
    If pnCol = 5 Then
        feComponente.TextMatrix(pnRow, pnCol) = IIf(Trim(Right(feComponente.TextMatrix(pnRow, 1), 2)) = "1", feComponente.TextMatrix(pnRow, pnCol), 0)
    End If
    If pnCol = 7 Or pnCol = 8 Or pnCol = 9 Then
        feComponente.TextMatrix(pnRow, pnCol) = UCase(feComponente.TextMatrix(pnRow, pnCol))
    End If
    If pnCol = 10 Then
        For i = 1 To feComponente.Rows - 1
            lnTotal = lnTotal + CCur(feComponente.TextMatrix(i, pnCol))
        Next
        txtTotal.Text = Format(lnTotal, gsFormatoNumeroView)
    End If
    feComponente.TextMatrix(pnRow, 5) = IIf(Trim(Right(feComponente.TextMatrix(pnRow, 1), 2)) = "1", Val(feComponente.TextMatrix(pnRow, 5)), 0)
End Sub
Private Sub feComponente_OnChangeCombo()
    Dim oAlmacen As New DLogAlmacen
    Dim lnTpoActivacion As Integer

    If feComponente.Col = 1 Then
        lnTpoActivacion = CInt(Trim(Right(feComponente.TextMatrix(feComponente.Row, 1), 2)))
        feComponente.TextMatrix(feComponente.Row, 2) = ""
        feComponente.TextMatrix(feComponente.Row, 4) = DevolverNuevoCodigoInventario(lnTpoActivacion, feComponente.Row)
        If lnTpoActivacion = 1 Then
            feComponente.rsTextBuscar = oAlmacen.GetBienesAlmacen(, gnLogBSTpoBienFijo)
        Else
            feComponente.rsTextBuscar = oAlmacen.GetBienesAlmacen(, gnLogBSTpoBienNoDepreciable)
        End If
        Call feComponente_OnEnterTextBuscar("", feComponente.Row, 2, False)
    End If
    Set oAlmacen = Nothing
End Sub
Private Sub feComponente_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    If pnCol = 1 Then
        If Len(Trim(feComponente.TextMatrix(pnRow, pnCol))) = 0 Then 'Si no ha escogido nada
            Cancel = False
        End If
    End If
    If pnCol = 3 Or pnCol = 4 Or pnCol = 6 Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
    End If
    If pnCol = 10 Then
        If Not IsNumeric(feComponente.TextMatrix(pnRow, pnCol)) Then
            MsgBox "Ingrese un Monto Decimal", vbInformation, "Aviso"
            Cancel = False
        End If
    End If
End Sub
Private Sub feComponente_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    Dim oBien As New DBien
    Dim lsTpoActivacion As String
    Dim lnBANCod As Integer
    
    lsTpoActivacion = Trim(Right(feComponente.TextMatrix(pnRow, 1), 2))
    If pnCol = 2 Then 'Codigo de Objeto
        If lsTpoActivacion = "1" And psDataCod <> "" Then
            lnBANCod = oBien.RecuperaCodigoBAN(psDataCod)
            If lnBANCod <> 0 Then
                feComponente.TextMatrix(pnRow, 6) = RecuperaMesesDepreciaTributariamente(lnBANCod, Mid(fsAgeCod, 4, 2))
            Else
                feComponente.TextMatrix(pnRow, 6) = 0
                MsgBox "El componente " & UCase(feComponente.TextMatrix(pnRow, 3)) & " No tiene configurado el Tipo de Activo Fijo", vbCritical, "Aviso"
            End If
        Else
            feComponente.TextMatrix(pnRow, 5) = 0
            feComponente.TextMatrix(pnRow, 6) = 0
        End If
    End If
    Set oBien = Nothing
End Sub
Private Function ValidarIngresoDatosFlex() As Boolean
    Dim i As Integer, J As Integer
    ValidarIngresoDatosFlex = True
    If Not FlexVacio(feComponente) Then 'Flex no esta Vacio
        For i = 1 To feComponente.Rows - 1
            For J = 1 To feComponente.Cols - 2
                If Len(Trim(feComponente.TextMatrix(i, J))) = 0 Then
                    ValidarIngresoDatosFlex = False
                    MsgBox "Ud. debe ingresar el dato " & UCase(feComponente.TextMatrix(0, J)), vbInformation, "Aviso"
                    feComponente.SetFocus
                    feComponente.Col = J
                    feComponente.Row = i
                    Exit Function
                End If
            Next
        Next
    End If
End Function
Private Sub Form_Unload(Cancel As Integer)
    fbAceptar = False
End Sub
Private Function DevolverNuevoCodigoInventario(ByVal pnTipoActivacion As Integer, ByVal pnFilaActual As Integer) As String
    Dim i As Integer
    Dim lsCodInventartioUltimo As String
    
    If pnTipoActivacion = 1 Then
        lsCodInventartioUltimo = fsCodInventarioUltimoAF
    Else
        lsCodInventartioUltimo = fsCodInventarioUltimoBND
    End If
    For i = 1 To feComponente.Rows - 1
        If Trim(Right(feComponente.TextMatrix(i, 1), 2)) <> "" Then
            If i <> pnFilaActual And pnTipoActivacion = CInt(Trim(Right(feComponente.TextMatrix(i, 1), 2))) Then
                feComponente.TextMatrix(i, 4) = Left(lsCodInventartioUltimo, Len(lsCodInventartioUltimo) - 5) & Format(CLng(Right(lsCodInventartioUltimo, 5)) + 1, "00000")
                lsCodInventartioUltimo = Trim(feComponente.TextMatrix(i, 4))
            End If
        End If
    Next
    DevolverNuevoCodigoInventario = Left(lsCodInventartioUltimo, Len(lsCodInventartioUltimo) - 5) & Format(CLng(Right(lsCodInventartioUltimo, 5)) + 1, "00000")
End Function
Private Function RecuperaMesesDepreciaTributariamente(ByVal pnTpoActivo As Integer, ByRef psAgeCod As String) As Integer
    Dim dLog As New DLogDeprecia
    Dim rs As New ADODB.Recordset
    If psAgeCod = "" Then psAgeCod = "01"
    Set rs = dLog.ObtienePorcentVidaUtlAF(pnTpoActivo)
    RecuperaMesesDepreciaTributariamente = 0
    Do While Not rs.EOF
        If psAgeCod = rs!cAgeCod Then
            RecuperaMesesDepreciaTributariamente = rs!nDepreMesT
            Exit Do
        End If
        rs.MoveNext
    Loop
    Set dLog = Nothing
    Set rs = Nothing
End Function
