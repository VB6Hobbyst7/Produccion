VERSION 5.00
Begin VB.Form frmNIIFNotasComplementariaConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notas Complementarias para Información Anual"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13665
   Icon            =   "frmNIIFNotasComplementariaConfig.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   13665
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   11400
      TabIndex        =   3
      Top             =   5205
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   12480
      TabIndex        =   2
      Top             =   5205
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selección de Items para reporte"
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
      Height          =   5025
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   13545
      Begin Sicmact.FlexEdit feNotas 
         Height          =   4725
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   13320
         _ExtentX        =   23495
         _ExtentY        =   8334
         Cols0           =   11
         HighLight       =   1
         AllowUserResizing=   1
         EncabezadosNombres=   "#-Id-Sel-Descripción de Notas-Sección-Columna-SeccionTmp-ColumnaTmp-Formula-Formula <= 2012-Aux"
         EncabezadosAnchos=   "0-0-500-4500-900-2800-0-0-2000-2000-0"
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
         ColumnasAEditar =   "X-X-2-X-4-5-X-X-8-9-X"
         ListaControles  =   "0-0-4-0-3-3-0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-L-L-L-L-C-C-L-L-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0"
         CantEntero      =   9
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         TipoBusqueda    =   0
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
      End
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   60
      TabIndex        =   4
      Top             =   5115
      Width           =   13545
   End
End
Attribute VB_Name = "frmNIIFNotasComplementariaConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************************
'** Nombre : frmNIIFNotasComplementariaConfig
'** Descripción : Configuración del Reporte Notas Complementarias Inf. Anual creado segun ERS149-2013
'** Creación : EJVG, 20140102 09:00:00 AM
'****************************************************************************************************
Option Explicit

Dim rsColumnaActivo As New ADODB.Recordset
Dim rsColumnaPasivo As New ADODB.Recordset

Private Sub feNotas_OnCellChange(pnRow As Long, pnCol As Long)
    If pnCol = 8 Or pnCol = 9 Then
        feNotas.TextMatrix(pnRow, pnCol) = Trim(feNotas.TextMatrix(pnRow, pnCol))
    End If
End Sub

Private Sub Form_Load()
    CargarVariables
    ListarConfiguracionNotas
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub CargarVariables()
    Dim oConst As New Dconstante
    Set rsColumnaActivo = oConst.RecuperaConstantes(10036)
    Set rsColumnaPasivo = oConst.RecuperaConstantes(10054)
    Set oConst = Nothing
End Sub
Private Sub ListarConfiguracionNotas()
    Dim oRep As New NRepFormula
    Dim rsNotas As New ADODB.Recordset

    On Error GoTo ErrListar
    Screen.MousePointer = 11
    Set rsNotas = oRep.RecuperaConfigRepNotasEstado(gContRepBaseNotasEstadoSitFinan)
    FormateaFlex feNotas
    If Not RSVacio(rsNotas) Then
        Do While Not rsNotas.EOF
            feNotas.AdicionaFila
            feNotas.TextMatrix(feNotas.row, 1) = rsNotas!nNotaEstado
            feNotas.TextMatrix(feNotas.row, 2) = IIf(rsNotas!bNotaComplementa, "1", "")
            feNotas.TextMatrix(feNotas.row, 3) = rsNotas!cDescripcion
            feNotas.TextMatrix(feNotas.row, 4) = IIf(rsNotas!bNotaComplementa, rsNotas!cSeccion & Space(75) & rsNotas!nSeccion, "")
            feNotas.TextMatrix(feNotas.row, 5) = IIf(rsNotas!bNotaComplementa, rsNotas!cColumna & Space(75) & rsNotas!nColumna, "")
            feNotas.TextMatrix(feNotas.row, 6) = feNotas.TextMatrix(feNotas.row, 4)
            feNotas.TextMatrix(feNotas.row, 7) = feNotas.TextMatrix(feNotas.row, 5)
            feNotas.TextMatrix(feNotas.row, 8) = rsNotas!cFormulaNotaComplementaria
            feNotas.TextMatrix(feNotas.row, 9) = rsNotas!cFormulaNotaComplementaria_2012
            rsNotas.MoveNext
        Loop
        feNotas.TopRow = 1
        feNotas.row = 1
    End If
    Set oRep = Nothing
    Set rsNotas = Nothing
    Screen.MousePointer = 0
    Exit Sub
ErrListar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub feNotas_RowColChange()
    Dim rsOpt As New ADODB.Recordset
    Dim lnSeccion As Integer
    If feNotas.Col = 4 Or feNotas.Col = 5 Or feNotas.Col = 8 Or feNotas.Col = 9 Then
        If feNotas.TextMatrix(feNotas.row, 2) <> "." Then
            MsgBox "Ud. debe primero debe marcar el Item", vbInformation, "Aviso"
            feNotas.Col = 2
            Exit Sub
        End If
        If feNotas.Col = 8 Or feNotas.Col = 9 Then Exit Sub
        With rsOpt
            .Fields.Append "desc", adVarChar, 1000
            .Fields.Append "value", adVarChar, 1
        End With
        If feNotas.Col = 4 Then
            With rsOpt
                .Open
                .AddNew
                .Fields("desc") = "Activos"
                .Fields("value") = "1"
                .AddNew
                .Fields("desc") = "Pasivos"
                .Fields("value") = "2"
            End With
        ElseIf feNotas.Col = 5 Then
            If feNotas.Col = 5 And Trim(feNotas.TextMatrix(feNotas.row, 4)) = "" Then
                MsgBox "Ud. debe primero seleccionar la sección", vbInformation, "Aviso"
                feNotas.Col = 4
                Exit Sub
            End If
            lnSeccion = CInt(Trim(Right(feNotas.TextMatrix(feNotas.row, 4), 2)))
            If lnSeccion = 1 Then
                rsColumnaActivo.MoveFirst
                rsOpt.Open
                Do While Not rsColumnaActivo.EOF
                    With rsOpt
                        .AddNew
                        .Fields("desc") = rsColumnaActivo!cConsDescripcion
                        .Fields("value") = rsColumnaActivo!nConsValor
                    End With
                    rsColumnaActivo.MoveNext
                Loop
            Else
                rsColumnaPasivo.MoveFirst
                rsOpt.Open
                Do While Not rsColumnaPasivo.EOF
                    With rsOpt
                        .AddNew
                        .Fields("desc") = rsColumnaPasivo!cConsDescripcion
                        .Fields("value") = rsColumnaPasivo!nConsValor
                    End With
                    rsColumnaPasivo.MoveNext
                Loop
            End If
        End If
        rsOpt.MoveFirst
        feNotas.CargaCombo rsOpt
    End If
    Set rsOpt = Nothing
End Sub
Private Sub feNotas_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    If feNotas.TextMatrix(pnRow, 2) = "." Then
        feNotas.TextMatrix(pnRow, 4) = feNotas.TextMatrix(pnRow, 6)
        feNotas.TextMatrix(pnRow, 5) = feNotas.TextMatrix(pnRow, 7)
    Else
        feNotas.TextMatrix(pnRow, 4) = ""
        feNotas.TextMatrix(pnRow, 5) = ""
    End If
End Sub
Private Sub feNotas_OnChangeCombo()
    If feNotas.TextMatrix(feNotas.row, 4) <> feNotas.TextMatrix(feNotas.row, 6) Then
        feNotas.TextMatrix(feNotas.row, 5) = ""
        feNotas.TextMatrix(feNotas.row, 6) = feNotas.TextMatrix(feNotas.row, 4)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("¿Esta seguro de salir la configuración de las Notas Complementarias?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Cancel = 1
        Exit Sub
    End If
    Set rsColumnaActivo = Nothing
    Set rsColumnaPasivo = Nothing
End Sub
Private Sub cmdGuardar_Click()
    Dim oNRep As NRepFormula
    Dim i As Integer, fila As Integer
    Dim Datos() As Variant
    Dim lbExito As Boolean
    
    On Error GoTo ErrGuardar
    If Not validaGuardar Then Exit Sub
    
    If MsgBox("¿Esta seguro de guardar la configuración de las Notas Complementarias?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    cmdGuardar.Enabled = False
    Set oNRep = New NRepFormula
    ReDim Datos(1 To 5, 0)
    For i = 1 To feNotas.Rows - 1
        If feNotas.TextMatrix(i, 2) = "." Then
            fila = fila + 1
            ReDim Preserve Datos(1 To 5, 0 To fila)
            Datos(1, fila) = CInt(feNotas.TextMatrix(i, 1))
            Datos(2, fila) = IIf(CInt(Trim(Right(feNotas.TextMatrix(i, 4), 2))) = 1, 1, 2) 'Seccion
            Datos(3, fila) = CInt(Trim(Right(feNotas.TextMatrix(i, 5), 2))) 'Columna
            Datos(4, fila) = Trim(feNotas.TextMatrix(i, 8)) 'Formula
            Datos(5, fila) = Trim(feNotas.TextMatrix(i, 9)) 'Formula_2012
        End If
    Next
    lbExito = oNRep.GuardarNotasComplementarias(gContRepBaseNotasEstadoSitFinan, Datos)
    Screen.MousePointer = 0
    cmdGuardar.Enabled = True
    If lbExito Then
        MsgBox "Se ha registrado las Notas Complementarias satisfactoriamente", vbInformation, "Aviso"
        ListarConfiguracionNotas
    Else
        MsgBox "Hubo un erro al guardar las Notas Complementarias, si el error persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If
    Exit Sub
ErrGuardar:
    Screen.MousePointer = 0
    cmdGuardar.Enabled = True
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub feNotas_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim Editar() As String
    Editar = Split(feNotas.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        Exit Sub
    End If
End Sub
Private Function validaGuardar() As Boolean
    Dim i As Integer
    Dim bSelecciono As Boolean
    validaGuardar = True
    If feNotas.TextMatrix(1, 0) = "" Then
        validaGuardar = False
        MsgBox "No existen datos para guardar", vbInformation, "Aviso"
        Exit Function
    End If
    For i = 1 To feNotas.Rows - 1
        If feNotas.TextMatrix(i, 2) = "." Then
            bSelecciono = True
            If Len(Trim(feNotas.TextMatrix(i, 4))) = 0 Then
                validaGuardar = False
                MsgBox "Ud. debe seleccionar la sección", vbInformation, "Aviso"
                feNotas.row = i
                feNotas.Col = 4
                feNotas.TopRow = i
                Exit Function
            End If
            If Len(Trim(feNotas.TextMatrix(i, 5))) = 0 Then
                validaGuardar = False
                MsgBox "Ud. debe seleccionar la columna", vbInformation, "Aviso"
                feNotas.row = i
                feNotas.Col = 5
                feNotas.TopRow = i
                Exit Function
            End If
            'Valida ingreso de Formulas
            If Len(Trim(feNotas.TextMatrix(i, 8))) > 0 Or Len(Trim(feNotas.TextMatrix(i, 9))) > 0 Then
                If Len(Trim(feNotas.TextMatrix(i, 8))) <= 0 Or Len(Trim(feNotas.TextMatrix(i, 9))) <= 0 Then
                    validaGuardar = False
                    MsgBox "Si va a ingresar formulas debe ingresar tanto en la columna [FORMULA] y [FORMULA <= 2012]", vbInformation, "Aviso"
                    feNotas.row = i
                    feNotas.Col = IIf(Len(Trim(feNotas.TextMatrix(i, 8))) <= 0, 8, 9)
                    feNotas.TopRow = i
                    Exit Function
                End If
            End If
        End If
    Next
    If Not bSelecciono Then
        validaGuardar = False
        MsgBox "No se ha seleccionado ningún registro, no se puede continuar..!", vbInformation, "Aviso"
        Exit Function
    End If
End Function
