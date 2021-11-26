VERSION 5.00
Begin VB.Form frmCredFormEval 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Formatos de Evaluación"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11985
   Icon            =   "frmCredFormEval.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
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
         TabIndex        =   5
         Top             =   5280
         Width           =   1170
      End
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
         Left            =   1440
         TabIndex        =   4
         Top             =   5280
         Width           =   1155
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
         Left            =   120
         TabIndex        =   3
         Top             =   5280
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
         Top             =   5280
         Width           =   1170
      End
      Begin SICMACT.FlexEdit feFormatosEval 
         Height          =   4935
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   8705
         Cols0           =   5
         HighLight       =   1
         EncabezadosNombres=   "-Formatos-Mínimo-Máximo-nCodForm"
         EncabezadosAnchos=   "300-5000-3000-3000-0"
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
         ColumnasAEditar =   "X-X-2-3-X"
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-R-C"
         FormatosEdit    =   "0-0-2-2-0"
         CantEntero      =   16
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
      End
   End
End
Attribute VB_Name = "frmCredFormEval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmCredFormEval
'** Descripción : Lista los formatos de Evaluacion con sus respectivos parámetros de mínimos y máximos
'**               Los mosntos serán de acuerdo al riesgo unico con el cuenta el cliente en el crédito
'**               evaluado, para asi poder determinar el formato correspondiente
'** Creación : PEAC, 20160520 13:51:01 PM
'**********************************************************************************************

Option Explicit

Private Sub cmdCancelar_Click()
 
    feFormatosEval.lbEditarFlex = False
    
    cmdEditar.Enabled = True
    cmdSalir.Enabled = True
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
    Call LlenaGrilla
    
End Sub

'Private Sub cmdAgregar_Click()
'    If ValidaRegistroRepetido Then
'        feFormatosEval.lbEditarFlex = True
'        feFormatosEval.AdicionaFila
'        feFormatosEval.SetFocus
'        SendKeys "{Enter}"
'        Call feFormatosEval_RowColChange
'    End If
'End Sub

Private Sub cmdEditar_Click()
'    cmdAgregar.Visible = True
'    cmdEliminar.Visible = True
    feFormatosEval.lbEditarFlex = True
    
    'feFormatosEval.BackColor = RGB(120, 125, 115)
    
    'cmdEditar.Visible = False
    cmdEditar.Enabled = False
    cmdSalir.Enabled = False
    cmdGrabar.Enabled = True
    cmdCancelar.Enabled = True
    
    'Call feFormatosEval_RowColChange
End Sub

Private Sub CmdEliminar_Click()
    If MsgBox("¿Está seguro de eliminar los datos de la fila " + CStr(feFormatosEval.row) + "?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        feFormatosEval.EliminaFila feFormatosEval.row
    End If
End Sub

Private Sub cmdGrabar_Click()
    Dim oNCred As COMDCredito.DCOMFormatosEval
    Dim i As Integer
    Dim nId As String
    Set oNCred = New COMDCredito.DCOMFormatosEval
    

    If ValidaRegistroRepetido Then
 '       If ValidaGrilla Then
            If MsgBox("Los Datos ingresados se guardarán, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub

            'Call oNCred.dEliminaCredEvalParametrosTipo
            
            For i = 1 To feFormatosEval.Rows - 1
                
'                If i < 10 Then
'                    nId = "000" & CStr(i)
'                ElseIf i < 100 Then
'                    nId = "00" & CStr(i)
'                ElseIf i < 1000 Then
'                    nId = "0" & CStr(i)
'                Else
'                    nId = CStr(i)
'                End If
'                nId = "PAR" & nId

                'Call oNCred.dInsertaCredEvalParametrosTipo(nId, Trim(Right(Trim(feFormatosEval.TextMatrix(i, 1)), 3)), Trim(Right(Trim(feFormatosEval.TextMatrix(i, 2)), 3)), feFormatosEval.TextMatrix(i, 3), feFormatosEval.TextMatrix(i, 4), Trim(Right(Trim(feFormatosEval.TextMatrix(i, 5)), 2)))
                Call oNCred.ActualizaConfFormatosEval(feFormatosEval.TextMatrix(i, 4), feFormatosEval.TextMatrix(i, 2), feFormatosEval.TextMatrix(i, 3))
            Next i
            
            feFormatosEval.lbEditarFlex = False
'            cmdAgregar.Visible = False
            'cmdEliminar.Visible = False
            cmdEditar.Visible = True
            
            MsgBox "Se realizaron los cambios satisfactoriamente.", vbInformation, "Atención"
            
    cmdEditar.Enabled = True
    cmdSalir.Enabled = True
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
            
            
 '       Else
 '           MsgBox "Faltan datos en la lista de parametros", vbInformation, "Aviso"
 '       End If
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

'Private Sub feFormatosEval_RowColChange()
'Dim oDCred As COMDCredito.DCOMCredito
'Dim oCons As COMDConstantes.DCOMConstantes
'   If feFormatosEval.lbEditarFlex Then
'        Set oDCred = New COMDCredito.DCOMCredito
'        Set oCons = New COMDConstantes.DCOMConstantes
'        Select Case feFormatosEval.col
'            Case 1 'Producto
'                feFormatosEval.CargaCombo oDCred.RecuperaProductosCrediticios
'            Case 2 'SubProducto
'                feFormatosEval.CargaCombo oDCred.RecuperaSubProductosCrediticios(Trim(Right(feFormatosEval.TextMatrix(feFormatosEval.Row, 1), 3)))
'            Case 5 'Tipo Evaluacion
'                feFormatosEval.CargaCombo oCons.RecuperaConstantes(7061)
'        End Select
'        Set oDCred = Nothing
'        Set oCons = Nothing
'    End If
'End Sub

Private Sub Form_Load()
        
    
Call LlenaGrilla
    
'    ''Dim oCred As COMDCredito.DCOMCredito
'    Dim oCred As COMDCredito.DCOMFormatosEval
'    Dim rsParam As ADODB.Recordset
'    Dim i As Integer
'    Set oCred = New COMDCredito.DCOMFormatosEval
'
'    Set rsParam = oCred.RecuperaConfigFornatosEval
'
'     If Not (rsParam.BOF And rsParam.EOF) Then
'        feFormatosEval.lbEditarFlex = True
'        Call LimpiaFlex(feFormatosEval)
'            For i = 0 To rsParam.RecordCount - 1
'                feFormatosEval.AdicionaFila
'                feFormatosEval.TextMatrix(i + 1, 0) = i + 1
'                feFormatosEval.TextMatrix(i + 1, 1) = rsParam!cNomFormato
'                feFormatosEval.TextMatrix(i + 1, 2) = Format(rsParam!nMontoMin, "#,##0.00")
'                feFormatosEval.TextMatrix(i + 1, 3) = Format(rsParam!nMontoMax, "#,##0.00")
'                feFormatosEval.TextMatrix(i + 1, 4) = rsParam!ncodForm
'                rsParam.MoveNext
'            Next i
'    End If
'    feFormatosEval.lbEditarFlex = False
End Sub

Public Sub LlenaGrilla()
    Dim oCred As COMDCredito.DCOMFormatosEval
    Dim rsParam As ADODB.Recordset
    Dim i As Integer
    Set oCred = New COMDCredito.DCOMFormatosEval
    
    Set rsParam = oCred.RecuperaConfigFornatosEval
    
     If Not (rsParam.BOF And rsParam.EOF) Then
        feFormatosEval.lbEditarFlex = True
        Call LimpiaFlex(feFormatosEval)
            For i = 0 To rsParam.RecordCount - 1
                feFormatosEval.AdicionaFila
                feFormatosEval.TextMatrix(i + 1, 0) = i + 1
                feFormatosEval.TextMatrix(i + 1, 1) = rsParam!cNomFormato
                feFormatosEval.TextMatrix(i + 1, 2) = Format(rsParam!nMontoMin, "#,##0.00")
                feFormatosEval.TextMatrix(i + 1, 3) = Format(rsParam!nMontoMax, "#,##0.00")
                feFormatosEval.TextMatrix(i + 1, 4) = rsParam!nCodForm
                rsParam.MoveNext
            Next i
    End If
    feFormatosEval.lbEditarFlex = False
End Sub
'    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
'    Dim R As ADODB.Recordset
'    Dim rs As ADODB.Recordset
'    '''Dim oPers As comdpersona.UCOMPersona
'    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
'    ''Set oPers = New comdpersona.UCOMPersona
'
'    Set R = oDCOMFormatosEval.RecuperaConfigFornatosEval
'    'Set rs = oDComPoliza.Garantias_x_Poliza(Trim(sNumPoliza))
'    Set oDCOMFormatosEval = Nothing
'
'    feFormatosEval.Clear
'    feFormatosEval.FormaCabecera
'    feFormatosEval.Rows = 2
'    feFormatosEval.rsFlex = R
'
''    feGarantias.Clear
''    feGarantias.FormaCabecera
''    feGarantias.Rows = 2
''    feGarantias.rsFlex = R
''    FePolizaGarant.rsFlex = rs
''    txtNumPoliza.Text = Trim(sNumPoliza)
''    Set oPers = Nothing
'End Sub

Public Function ValidaRegistroRepetido() As Boolean
    Dim i As Integer, J As Integer, UltRegistro As Integer
    ValidaRegistroRepetido = False
    UltRegistro = feFormatosEval.Rows - 1
    For i = 1 To feFormatosEval.Rows - 1
        If feFormatosEval.TextMatrix(i, 0) <> "" Then
            For J = 1 To feFormatosEval.Rows - 1
                If i <> J Then
                    If Trim(feFormatosEval.TextMatrix(i, 1)) = Trim(feFormatosEval.TextMatrix(J, 1)) And _
                    Trim(feFormatosEval.TextMatrix(i, 2)) = Trim(feFormatosEval.TextMatrix(J, 2)) And _
                    Trim(feFormatosEval.TextMatrix(i, 3)) = Trim(feFormatosEval.TextMatrix(J, 3)) Then
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
    For i = 1 To feFormatosEval.Rows - 1
        If feFormatosEval.TextMatrix(i, 0) <> "" Then
            If Trim(feFormatosEval.TextMatrix(i, 1)) = "" Or Trim(feFormatosEval.TextMatrix(i, 2)) = "" Or _
               Trim(feFormatosEval.TextMatrix(i, 3)) = "" Or Trim(feFormatosEval.TextMatrix(i, 4)) = "" Or _
               Trim(feFormatosEval.TextMatrix(i, 5)) = "" Then
                ValidaGrilla = False
                Exit Function
            End If
        End If
    Next i
    ValidaGrilla = True
End Function

