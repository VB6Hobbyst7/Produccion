VERSION 5.00
Begin VB.Form frmCredEvalParamTipos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parámetros de Tipo de Evaluación"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   Icon            =   "frmCredEvalParamTipos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar"
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
         TabIndex        =   6
         Top             =   3000
         Width           =   1155
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
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
         TabIndex        =   5
         Top             =   3000
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar"
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
         TabIndex        =   4
         Top             =   3000
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
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
         Left            =   7680
         TabIndex        =   3
         Top             =   3000
         Width           =   1170
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
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
         Left            =   8880
         TabIndex        =   2
         Top             =   3000
         Width           =   1170
      End
      Begin SICMACT.FlexEdit fgTipoEval 
         Height          =   2655
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   4683
         Cols0           =   7
         HighLight       =   1
         EncabezadosNombres=   "-Producto-Sub Producto-Min S/.-Max S/.-Tipo Evaluación-Aux"
         EncabezadosAnchos=   "300-2500-2500-1400-1400-1500-0"
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
         ColumnasAEditar =   "X-1-2-3-4-5-X"
         ListaControles  =   "0-3-3-0-0-3-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-R-R-L-C"
         FormatosEdit    =   "0-0-0-2-2-0-0"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
      End
   End
End
Attribute VB_Name = "frmCredEvalParamTipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmCredEvalParamTipos
'** Descripción : Administración de parametros de los tipo de evaluación de crédito creado
'**               segun RFC090-2012
'** Creación : JUEZ, 20120903 09:00:00 AM
'**********************************************************************************************

Option Explicit

Private Sub cmdAgregar_Click()
    If ValidaRegistroRepetido Then
        fgTipoEval.lbEditarFlex = True
        fgTipoEval.AdicionaFila
        fgTipoEval.SetFocus
        SendKeys "{Enter}"
        Call fgTipoEval_RowColChange
    End If
End Sub

Private Sub CmdEditar_Click()
    cmdAgregar.Visible = True
    cmdEliminar.Visible = True
    fgTipoEval.lbEditarFlex = True
    cmdEditar.Visible = False
    Call fgTipoEval_RowColChange
End Sub

Private Sub CmdEliminar_Click()
    If MsgBox("¿Está seguro de eliminar los datos de la fila " + CStr(fgTipoEval.Row) + "?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        fgTipoEval.EliminaFila fgTipoEval.Row
    End If
End Sub

Private Sub CmdGrabar_Click()
    Dim oNCred As COMDCredito.DCOMCredActBD
    Dim i As Integer
    Dim nId As String
    Set oNCred = New COMDCredito.DCOMCredActBD
    
    If ValidaRegistroRepetido Then
        If ValidaGrilla Then
            If MsgBox("Los Datos seran Grabados, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
            
            Call oNCred.dEliminaCredEvalParametrosTipo
            For i = 1 To fgTipoEval.Rows - 1
                If i < 10 Then
                    nId = "000" & CStr(i)
                ElseIf i < 100 Then
                    nId = "00" & CStr(i)
                ElseIf i < 1000 Then
                    nId = "0" & CStr(i)
                Else
                    nId = CStr(i)
                End If
                nId = "PAR" & nId
                Call oNCred.dInsertaCredEvalParametrosTipo(nId, Trim(Right(Trim(fgTipoEval.TextMatrix(i, 1)), 3)), Trim(Right(Trim(fgTipoEval.TextMatrix(i, 2)), 3)), fgTipoEval.TextMatrix(i, 3), fgTipoEval.TextMatrix(i, 4), Trim(Right(Trim(fgTipoEval.TextMatrix(i, 5)), 2)))
            Next i
            fgTipoEval.lbEditarFlex = False
            cmdAgregar.Visible = False
            cmdEliminar.Visible = False
            cmdEditar.Visible = True
            MsgBox "Los parámetros se grabaron con exito", vbInformation, "Aviso"
        Else
            MsgBox "Faltan datos en la lista de parametros", vbInformation, "Aviso"
        End If
    End If
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub fgTipoEval_RowColChange()
Dim oDCred As COMDCredito.DCOMCredito
Dim oCons As COMDConstantes.DCOMConstantes
   If fgTipoEval.lbEditarFlex Then
        Set oDCred = New COMDCredito.DCOMCredito
        Set oCons = New COMDConstantes.DCOMConstantes
        Select Case fgTipoEval.Col
            Case 1 'Producto
                fgTipoEval.CargaCombo oDCred.RecuperaProductosCrediticios
            Case 2 'SubProducto
                fgTipoEval.CargaCombo oDCred.RecuperaSubProductosCrediticios(Trim(Right(fgTipoEval.TextMatrix(fgTipoEval.Row, 1), 3)))
            Case 5 'Tipo Evaluacion
                fgTipoEval.CargaCombo oCons.RecuperaConstantes(7061)
        End Select
        Set oDCred = Nothing
        Set oCons = Nothing
    End If
End Sub

Private Sub Form_Load()
    Dim oCred As COMDCredito.DCOMCredito
    Dim rsParam As ADODB.Recordset
    Dim i As Integer
    Set oCred = New COMDCredito.DCOMCredito
    
    Set rsParam = oCred.ListaCredEvalParamTipos
    
     If Not (rsParam.BOF And rsParam.EOF) Then
        fgTipoEval.lbEditarFlex = True
        Call LimpiaFlex(fgTipoEval)
            For i = 0 To rsParam.RecordCount - 1
                fgTipoEval.AdicionaFila
                fgTipoEval.TextMatrix(i + 1, 0) = i + 1
                fgTipoEval.TextMatrix(i + 1, 1) = rsParam!cTpoProdDesc & Space(75) & rsParam!cTpoProdCod
                fgTipoEval.TextMatrix(i + 1, 2) = rsParam!cTpoSubProdDesc & Space(75) & rsParam!cTpoSubProdCod
                fgTipoEval.TextMatrix(i + 1, 3) = Format(rsParam!nMin, "#,##0.00")
                fgTipoEval.TextMatrix(i + 1, 4) = Format(rsParam!nMax, "#,##0.00")
                fgTipoEval.TextMatrix(i + 1, 5) = rsParam!cTpoEvalDesc & Space(75) & rsParam!nTpoEvalCod
                rsParam.MoveNext
            Next i
    End If
    fgTipoEval.lbEditarFlex = False
End Sub

Public Function ValidaRegistroRepetido() As Boolean
    Dim i As Integer, J As Integer, UltRegistro As Integer
    ValidaRegistroRepetido = False
    UltRegistro = fgTipoEval.Rows - 1
    For i = 1 To fgTipoEval.Rows - 1
        If fgTipoEval.TextMatrix(i, 0) <> "" Then
            For J = 1 To fgTipoEval.Rows - 1
                If i <> J Then
                    If Trim(fgTipoEval.TextMatrix(i, 1)) = Trim(fgTipoEval.TextMatrix(J, 1)) And _
                    Trim(fgTipoEval.TextMatrix(i, 2)) = Trim(fgTipoEval.TextMatrix(J, 2)) And _
                    Trim(fgTipoEval.TextMatrix(i, 3)) = Trim(fgTipoEval.TextMatrix(J, 3)) And _
                    Trim(fgTipoEval.TextMatrix(i, 4)) = Trim(fgTipoEval.TextMatrix(J, 4)) And _
                    Trim(fgTipoEval.TextMatrix(i, 5)) = Trim(fgTipoEval.TextMatrix(J, 5)) Then
                        ValidaRegistroRepetido = False
                        MsgBox "El registro de la fila " & J & " ya existe en la fila " & i & ", favor de modificar", vbInformation, "Aviso"
                        Exit Function
                    End If
                End If
            Next J
        End If
    Next i
    ValidaRegistroRepetido = True
End Function

Public Function ValidaGrilla() As Boolean
    Dim i As Integer
    
    ValidaGrilla = False
    For i = 1 To fgTipoEval.Rows - 1
        If fgTipoEval.TextMatrix(i, 0) <> "" Then
            If Trim(fgTipoEval.TextMatrix(i, 1)) = "" Or Trim(fgTipoEval.TextMatrix(i, 2)) = "" Or _
               Trim(fgTipoEval.TextMatrix(i, 3)) = "" Or Trim(fgTipoEval.TextMatrix(i, 4)) = "" Or _
               Trim(fgTipoEval.TextMatrix(i, 5)) = "" Then
                ValidaGrilla = False
                Exit Function
            End If
        End If
    Next i
    ValidaGrilla = True
End Function
