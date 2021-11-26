VERSION 5.00
Begin VB.Form frmCredBPPParametros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parámetros BPP"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12135
   Icon            =   "frmCredBPPParametros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
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
      Left            =   10800
      TabIndex        =   15
      Top             =   7680
      Width           =   1170
   End
   Begin VB.CommandButton cmdEditarParam 
      Caption         =   "Editar"
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
      Left            =   120
      TabIndex        =   13
      Top             =   7680
      Width           =   1170
   End
   Begin VB.CommandButton cmdEliminarParam 
      Caption         =   "Eliminar"
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
      Left            =   1320
      TabIndex        =   14
      Top             =   7680
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Caption         =   " Registro "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   11895
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
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
         Left            =   9360
         TabIndex        =   11
         Top             =   4440
         Width           =   1170
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
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
         Left            =   10560
         TabIndex        =   12
         Top             =   4440
         Width           =   1170
      End
      Begin VB.Frame Frame5 
         Caption         =   " Categoria y Tamaño de Cartera "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   21
         Top             =   3480
         Width           =   11655
         Begin VB.ComboBox cboCategoriaAna 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   320
            Width           =   2055
         End
         Begin SICMACT.EditMoney txtMinCartera 
            Height          =   300
            Left            =   7800
            TabIndex        =   9
            Top             =   315
            Width           =   1575
            _ExtentX        =   2778
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
         Begin SICMACT.EditMoney txtMaxCartera 
            Height          =   300
            Left            =   9840
            TabIndex        =   10
            Top             =   320
            Width           =   1575
            _ExtentX        =   2778
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
         Begin VB.Label Label4 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9600
            TabIndex        =   24
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label3 
            Caption         =   "Tamaño de Cartera S/. :"
            Height          =   255
            Left            =   5760
            TabIndex        =   23
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Categoria de Analista :"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " Indicadores de Cartera atrasada  "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   7200
         TabIndex        =   20
         Top             =   720
         Width           =   4575
         Begin VB.CommandButton cmdEliminarInd 
            Caption         =   "Eliminar"
            Height          =   345
            Left            =   1200
            TabIndex        =   7
            Top             =   1920
            Width           =   1050
         End
         Begin VB.CommandButton cmdAgregarInd 
            Caption         =   "Agregar"
            Height          =   345
            Left            =   120
            TabIndex        =   6
            Top             =   1920
            Width           =   1050
         End
         Begin SICMACT.FlexEdit feIndCartAtrasada 
            Height          =   1575
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   4320
            _ExtentX        =   7620
            _ExtentY        =   2778
            Cols0           =   5
            HighLight       =   1
            EncabezadosNombres=   "-Desde (dias)-Hasta (dias)-+ Judiciales-Aux"
            EncabezadosAnchos=   "300-1150-1150-1300-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-2-3-X"
            ListaControles  =   "0-0-0-4-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-R-R-L-C"
            FormatosEdit    =   "0-3-3-0-0"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            ColWidth0       =   300
            RowHeight0      =   300
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Aplicable a Agencias "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   3600
         TabIndex        =   19
         Top             =   720
         Width           =   3495
         Begin VB.ListBox lstAgencias 
            Height          =   1635
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   4
            Top             =   550
            Width           =   3255
         End
         Begin VB.CheckBox chkTodosAgencia 
            Caption         =   "Todos"
            Height          =   255
            Left            =   160
            TabIndex        =   3
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Aplicable a Sub Productos "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   3375
         Begin VB.CheckBox chkTodosSubProductos 
            Caption         =   "Todos"
            Height          =   255
            Left            =   160
            TabIndex        =   1
            Top             =   240
            Width           =   1215
         End
         Begin VB.ListBox lstSubProd 
            Height          =   1635
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   2
            Top             =   555
            Width           =   3135
         End
      End
      Begin VB.ComboBox cboTipoCartera 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   320
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Cartera :"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
   End
   Begin SICMACT.FlexEdit feParametrosBPP 
      Height          =   2415
      Left            =   120
      TabIndex        =   25
      Top             =   5160
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   4260
      Cols0           =   9
      HighLight       =   1
      EncabezadosNombres=   "-Categoría-Tipo de Cartera-Min Cartera-Max Cartera-SubProdts.-Agencias-Mora-cIdParametro"
      EncabezadosAnchos=   "300-1600-3200-1400-1400-1200-1200-1200-0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-R-R-C-C-C-L"
      FormatosEdit    =   "0-1-1-2-2-0-0-0-0"
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   3
      ColWidth0       =   300
      RowHeight0      =   300
   End
End
Attribute VB_Name = "frmCredBPPParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmCreBPPParametros
'** Descripción : Formulario para la Administracion de los parametros para el BPP
'**               creado segun RFC099-2012
'** Creación : JUEZ, 20121012 09:00:00 AM
'**********************************************************************************************

Option Explicit
Dim fbNuevo As Boolean
Dim fbActualiza As Boolean

Private Sub cboCategoriaAna_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMinCartera.SetFocus
    End If
End Sub

Private Sub cboTipoCartera_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lstSubProd.SetFocus
    End If
End Sub

Private Sub chkTodosAgencia_Click()
    Call CheckLista(IIf(chkTodosAgencia.value = 1, True, False), lstAgencias)
End Sub

Private Sub chkTodosSubProductos_Click()
    Call CheckLista(IIf(chkTodosSubProductos.value = 1, True, False), lstSubProd)
End Sub

Private Sub CheckLista(ByVal bCheck As Boolean, ByVal lstLista As ListBox)
    Dim i As Integer
    For i = 0 To lstLista.ListCount - 1
        lstLista.Selected(i) = bCheck
    Next i
    'lstLista.Enabled = IIf(bCheck, False, True)
End Sub

Private Sub cmdAgregarInd_Click()
    feIndCartAtrasada.AdicionaFila
    feIndCartAtrasada.SetFocus
    SendKeys "{Enter}"
End Sub

Private Sub cmdCancelar_Click()
    'If MsgBox("¿Seguro de cancelar el registro?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        LimpiaControles
        fbNuevo = True
        fbActualiza = False
        cmdEditarParam.Enabled = True
        cmdEliminarParam.Enabled = True
        feParametrosBPP.Enabled = True
        CargaDatosParametros
    'End If
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdEditarParam_Click()
    Dim oLista As COMDCredito.DCOMBPPR
    Dim rs As ADODB.Recordset
    Dim lnFila As Integer
    
    fbActualiza = True
    fbNuevo = False
    cmdEditarParam.Enabled = False
    cmdEliminarParam.Enabled = False
    feParametrosBPP.Enabled = False
    
    Set oLista = New COMDCredito.DCOMBPPR
    Set rs = oLista.RecuperaParametroDatos(feParametrosBPP.TextMatrix(feParametrosBPP.Row, 8))
    Set oLista = Nothing
    
    cboTipoCartera.ListIndex = IndiceListaCombo(cboTipoCartera, rs!cIdTipoCartera)
    cboCategoriaAna.ListIndex = IndiceListaCombo(cboCategoriaAna, rs!cIdCatAnalista)
    txtMinCartera.Text = Format(rs!nMinCart, "#,##0.00")
    txtMaxCartera.Text = Format(rs!nMaxCart, "#,##0.00")
    
    lstSubProd.Clear
    Call LlenaListas(lstSubProd, 1)
    lstAgencias.Clear
    Call LlenaListas(lstAgencias, 2)
    
    Set oLista = New COMDCredito.DCOMBPPR
    Set rs = oLista.RecuperaParametrosLista(feParametrosBPP.TextMatrix(feParametrosBPP.Row, 8), 3)
    Set oLista = Nothing
    
    Call LimpiaFlex(feIndCartAtrasada)
    Do While Not rs.EOF
        feIndCartAtrasada.AdicionaFila
        lnFila = feIndCartAtrasada.Row
        feIndCartAtrasada.TextMatrix(lnFila, 1) = rs!nDiasDesde
        feIndCartAtrasada.TextMatrix(lnFila, 2) = rs!nDiasHasta
        feIndCartAtrasada.TextMatrix(lnFila, 3) = IIf(rs!nJudicial = 1, 1, "")
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Private Sub cmdEliminarInd_Click()
    If MsgBox("¿Está seguro de eliminar los datos de la fila " + CStr(feIndCartAtrasada.Row) + "?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        feIndCartAtrasada.EliminaFila feIndCartAtrasada.Row
    End If
End Sub

Private Sub cmdEliminarParam_Click()
    If feParametrosBPP.TextMatrix(feParametrosBPP.Row, 0) <> "" Then
        If MsgBox("¿Está seguro de eliminar los datos de la fila " + CStr(feParametrosBPP.Row) + "?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            Dim oCredBPP As COMNCredito.NCOMBPPR
            Set oCredBPP = New COMNCredito.NCOMBPPR
            Call oCredBPP.dEliminaParametro(feParametrosBPP.TextMatrix(feParametrosBPP.Row, 8))
            feParametrosBPP.EliminaFila feParametrosBPP.Row
            CargaDatosParametros
        End If
    End If
End Sub

Private Sub CmdGrabar_Click()
    If ValidaDatos Then
        Dim oCredBPP As COMNCredito.NCOMBPPR
        Dim MatTpoProd() As String
        Dim MatAgencias() As String
        Dim MatIndicad() As String
        Dim cCategoriaPertenece As String
        
        ReDim MatTpoProd(DevuelveCantidadCheckList(lstSubProd), 1)
        ReDim MatAgencias(DevuelveCantidadCheckList(lstAgencias), 1)
        ReDim MatIndicad(feIndCartAtrasada.Rows - 1, 3)
        
        MatTpoProd = LlenaMatriz(lstSubProd, 1)
        MatAgencias = LlenaMatriz(lstAgencias, 1)
        
        MatIndicad = LlenaMatriz(lstAgencias, 3)
        
        Set oCredBPP = New COMNCredito.NCOMBPPR

        If oCredBPP.VerificaSiExisteParametro(Trim(Right(cboTipoCartera.Text, 10)), Trim(Right(cboCategoriaAna.Text, 10)), IIf(fbNuevo, "", feParametrosBPP.TextMatrix(feParametrosBPP.Row, 8))) = False Then
            If oCredBPP.VerificaTamanoCartera(IIf(fbNuevo, "", feParametrosBPP.TextMatrix(feParametrosBPP.Row, 8)), Trim(Right(cboTipoCartera.Text, 10)), _
                                                txtMinCartera.Text, txtMaxCartera.Text, cCategoriaPertenece) = False Then
                If fbNuevo Then
                    If MsgBox("¿Está seguro de " + IIf(fbNuevo = True, "registrar", "actualizar") + " los datos?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
                    
                    Call oCredBPP.dInsertaParametro(Trim(Right(cboTipoCartera.Text, 10)), MatTpoProd, MatAgencias, MatIndicad, _
                                                    Trim(Right(cboCategoriaAna.Text, 10)), txtMinCartera.Text, txtMaxCartera.Text)
                    MsgBox "Los datos se registraron correctamente", vbInformation, "Aviso"
                    Call cmdCancelar_Click
                    CargaDatosParametros
                ElseIf fbActualiza Then
                    If MsgBox("¿Está seguro de " + IIf(fbNuevo = True, "registrar", "actualizar") + " los datos?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
                    
                    Call oCredBPP.dActualizaParametro(Trim(Right(cboTipoCartera.Text, 10)), MatTpoProd, MatAgencias, MatIndicad, _
                                                    Trim(Right(cboCategoriaAna.Text, 10)), txtMinCartera.Text, txtMaxCartera.Text, _
                                                    feParametrosBPP.TextMatrix(feParametrosBPP.Row, 8))
                    MsgBox "Los datos se actualizaron correctamente", vbInformation, "Aviso"
                    Call cmdCancelar_Click
                    CargaDatosParametros
                End If
            Else
                MsgBox "El tamaño de la cartera ingresada se cruza con la categoria " + cCategoriaPertenece + " del mismo tipo de cartera, Favor de verificar", vbExclamation, "Aviso!"
                Exit Sub
            End If
        Else
            MsgBox "El tipo de Cartera y la Categoria coinciden con uno de los parametros ya registrados, Favor de verificar", vbExclamation, "Aviso!"
            Exit Sub
        End If
    End If
End Sub

Private Sub feIndCartAtrasada_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then
    '    If feIndCartAtrasada.Col = 2 Then
    '        If CLng(feIndCartAtrasada.TextMatrix(feIndCartAtrasada.Row, 2)) > 9999999 Then
    '            feIndCartAtrasada.TextMatrix(feIndCartAtrasada.Row, 2) = Mid(feIndCartAtrasada.TextMatrix(feIndCartAtrasada.Row, 2), 1, Len(feIndCartAtrasada.TextMatrix(feIndCartAtrasada.Row, 2)) - 1)
    '        End If
    '    End If
    'End If
End Sub

Private Sub feIndCartAtrasada_OnCellChange(pnRow As Long, pnCol As Long)
    If Len(CStr(Replace(feIndCartAtrasada.TextMatrix(pnRow, pnCol), ",", ""))) > 7 Then
        feIndCartAtrasada.TextMatrix(pnRow, pnCol) = Mid(Replace(feIndCartAtrasada.TextMatrix(pnRow, pnCol), ",", ""), 1, 7)
        feIndCartAtrasada.TextMatrix(pnRow, pnCol) = Format(feIndCartAtrasada.TextMatrix(pnRow, pnCol), "#,##0")
    End If
    If feIndCartAtrasada.TextMatrix(pnRow, 2) <> "" And feIndCartAtrasada.TextMatrix(pnRow, 1) <> "" Then
        If CLng(feIndCartAtrasada.TextMatrix(pnRow, 2)) <= CLng(feIndCartAtrasada.TextMatrix(pnRow, 1)) Then
            MsgBox "El valor Hasta debe ser mayor que el valor Desde, favor de corregir", vbExclamation, "Aviso"
            feIndCartAtrasada.TextMatrix(pnRow, 2) = ""
            feIndCartAtrasada.TextMatrix(pnRow, 1) = ""
        End If
    End If
End Sub

Private Sub feParametrosBPP_Click()
    If feParametrosBPP.TextMatrix(feParametrosBPP.Row, feParametrosBPP.Col) <> "" Then
        If feParametrosBPP.Col = 5 Then
            frmCredBPPLista.Inicio feParametrosBPP.TextMatrix(feParametrosBPP.Row, 8), 1
        ElseIf feParametrosBPP.Col = 6 Then
            frmCredBPPLista.Inicio feParametrosBPP.TextMatrix(feParametrosBPP.Row, 8), 2
        ElseIf feParametrosBPP.Col = 7 Then
            frmCredBPPLista.Inicio feParametrosBPP.TextMatrix(feParametrosBPP.Row, 8), 3
        End If
    End If
End Sub

Private Sub feParametrosBPP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If feParametrosBPP.TextMatrix(feParametrosBPP.Row, feParametrosBPP.Col) <> "" Then
            If feParametrosBPP.Col = 5 Then
                frmCredBPPLista.Inicio feParametrosBPP.TextMatrix(feParametrosBPP.Row, 8), 1
            ElseIf feParametrosBPP.Col = 6 Then
                frmCredBPPLista.Inicio feParametrosBPP.TextMatrix(feParametrosBPP.Row, 8), 2
            ElseIf feParametrosBPP.Col = 7 Then
                frmCredBPPLista.Inicio feParametrosBPP.TextMatrix(feParametrosBPP.Row, 8), 3
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    fbNuevo = True
    fbActualiza = False
    Call ListarTiposCartera
    Call ListarAgencias
    Call ListarSubProductos
    Call ListarCategoriaAnalistas
    
    CargaDatosParametros
End Sub

Private Sub ListarTiposCartera()
    Dim oCart As COMNCredito.NCOMBPPR
    Dim rsCart As ADODB.Recordset
    
    Set oCart = New COMNCredito.NCOMBPPR
    Set rsCart = oCart.ListarTiposCartera()
    Set oCart = Nothing
    cboTipoCartera.Clear
    While Not rsCart.EOF
        cboTipoCartera.AddItem rsCart.fields(1) & Space(500) & rsCart.fields(0)
        rsCart.MoveNext
    Wend
    Set rsCart = Nothing
End Sub

Private Sub ListarAgencias()
    Dim oAge As COMDConstantes.DCOMAgencias
    Dim rsAgencias As ADODB.Recordset
    Set oAge = New COMDConstantes.DCOMAgencias
        Set rsAgencias = oAge.ObtieneAgencias()
    Set oAge = Nothing
    If rsAgencias Is Nothing Then
        MsgBox " No se encuentran las Agencias ", vbInformation, " Aviso "
    Else
        lstAgencias.Clear
        With rsAgencias
            Do While Not rsAgencias.EOF
                lstAgencias.AddItem rsAgencias!nConsValor & " " & Trim(rsAgencias!cConsDescripcion)
                rsAgencias.MoveNext
            Loop
        End With
        lstAgencias.Selected(0) = True
    End If
End Sub

Private Sub ListarSubProductos()
    Dim oSubProd As COMDConstantes.DCOMConstantes
    Dim rsSubProd As ADODB.Recordset
    Set oSubProd = New COMDConstantes.DCOMConstantes
        Set rsSubProd = oSubProd.RecuperaConstantes(3033, 1)
    Set oSubProd = Nothing
    If rsSubProd Is Nothing Then
        MsgBox " No se encuentran las Agencias ", vbInformation, " Aviso "
    Else
        lstSubProd.Clear
        With rsSubProd
            Do While Not rsSubProd.EOF
                lstSubProd.AddItem rsSubProd!nConsValor & " " & Trim(rsSubProd!cConsDescripcion)
                
                rsSubProd.MoveNext
            Loop
        End With
        lstSubProd.Selected(0) = True
    End If
End Sub

Private Sub ListarCategoriaAnalistas()
    Dim oCat As COMNCredito.NCOMBPPR
    Dim rsCat As ADODB.Recordset
    
    Set oCat = New COMNCredito.NCOMBPPR
    Set rsCat = oCat.ListarCategoriaAnalistas()
    
    cboCategoriaAna.Clear
    While Not rsCat.EOF
        cboCategoriaAna.AddItem rsCat.fields(1) & Space(100) & rsCat.fields(0)
        rsCat.MoveNext
    Wend
    Set rsCat = Nothing
End Sub

Private Sub txtMaxCartera_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGrabar.SetFocus
    End If
End Sub

Private Sub txtMinCartera_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMaxCartera.SetFocus
    End If
End Sub

Private Sub LimpiaControles()
    ListarTiposCartera
    chkTodosAgencia.value = 0
    ListarAgencias
    chkTodosSubProductos.value = 0
    ListarSubProductos
    ListarCategoriaAnalistas
    txtMinCartera.Text = 0
    txtMaxCartera.Text = 0
    Call LimpiaFlex(feIndCartAtrasada)
End Sub

Private Sub CargaDatosParametros()
    Dim oCredBPP As COMDCredito.DCOMBPPR
    Dim rs As ADODB.Recordset
    Dim lnFila As Integer
    Set oCredBPP = New COMDCredito.DCOMBPPR
    
    Set rs = oCredBPP.RecuperaParametroDatos()
    Set oCredBPP = Nothing
    Call LimpiaFlex(feParametrosBPP)
    If Not rs.EOF Then
        Do While Not rs.EOF
            feParametrosBPP.AdicionaFila
            lnFila = feParametrosBPP.Row
            feParametrosBPP.TextMatrix(lnFila, 1) = rs!cCategoria
            feParametrosBPP.TextMatrix(lnFila, 2) = rs!cTipoCartera
            feParametrosBPP.TextMatrix(lnFila, 3) = Format(rs!nMinCart, "#,##0.00")
            feParametrosBPP.TextMatrix(lnFila, 4) = Format(rs!nMaxCart, "#,##0.00")
            feParametrosBPP.TextMatrix(lnFila, 5) = "Ver"
            feParametrosBPP.TextMatrix(lnFila, 6) = "Ver"
            feParametrosBPP.TextMatrix(lnFila, 7) = "Ver"
            feParametrosBPP.TextMatrix(lnFila, 8) = rs!cIdParametro
            rs.MoveNext
        Loop
    Else
        cmdEditarParam.Enabled = False
        cmdEliminarParam.Enabled = False
        feParametrosBPP.Enabled = False
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Function ValidaDatos() As Boolean
    Dim i As Integer
    Dim CTpoProd As Integer
    Dim CAgencia As Integer
    ValidaDatos = False
    
    CTpoProd = DevuelveCantidadCheckList(lstSubProd)
    CAgencia = DevuelveCantidadCheckList(lstAgencias)
    
    If cboTipoCartera.Text = "" Then
        MsgBox "Debe seleccionar el tipo de Cartera", vbInformation, "Aviso"
        cboTipoCartera.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If CTpoProd = 0 Then
        MsgBox "Debe seleccionar al menos un Tipo de Producto", vbInformation, "Aviso"
        lstSubProd.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If CAgencia = 0 Then
        MsgBox "Debe seleccionar al menos una Agencia", vbInformation, "Aviso"
        lstAgencias.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If ValidaGrillas(feIndCartAtrasada) = False Then
        ValidaDatos = False
        Exit Function
    End If
    If cboCategoriaAna.Text = "" Then
        MsgBox "Debe seleccionar la categoria del analista", vbInformation, "Aviso"
        cboCategoriaAna.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If CDbl(txtMaxCartera.Text) <= 0 Then
        MsgBox "Debe ingresar el tamaño maximo de la cartera", vbInformation, "Aviso"
        txtMaxCartera.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If CDbl(txtMaxCartera.Text) < CDbl(txtMinCartera.Text) Then
        MsgBox "El tamaño maximo de la cartera debe ser mayor al tamaño minimo", vbInformation, "Aviso"
        txtMaxCartera.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    ValidaDatos = True
End Function

Private Function DevuelveCantidadCheckList(ByVal lstLista As ListBox) As Integer
    Dim i As Integer
    Dim Cant As Integer
    
    For i = 1 To lstLista.ListCount
        If lstLista.Selected(i - 1) = True Then
            Cant = Cant + 1
        End If
    Next
    DevuelveCantidadCheckList = Cant
End Function

Private Function ValidaGrillas(ByVal Flex As FlexEdit) As Boolean
    Dim i As Integer
    ValidaGrillas = False
    If Flex.Rows - 1 = 1 And Flex.TextMatrix(1, 0) = "" Then
        MsgBox "Debe ingresar los indicadores de cartera atrasada", vbInformation, "Aviso"
        cmdAgregarInd.SetFocus
        ValidaGrillas = False
        Exit Function
    End If
    
    For i = 1 To Flex.Rows - 1
        If Flex.TextMatrix(i, 0) <> "" Then
            If Trim(Flex.TextMatrix(i, 1)) = "" Or Trim(Flex.TextMatrix(i, 2)) = "" Then
                MsgBox "Faltan datos en los indicadores de cartera atrasada", vbInformation, "Aviso"
                cmdAgregarInd.SetFocus
                ValidaGrillas = False
                Exit Function
            End If
        End If
    Next i
    ValidaGrillas = True
End Function

Private Function LlenaMatriz(ByVal lstLista As ListBox, ByVal pnValCant As Integer) As Variant
    Dim MatLista() As String
    Dim i As Integer
    Dim nTamano As Integer
    If pnValCant = 1 Then
        ReDim MatLista(DevuelveCantidadCheckList(lstLista), 1)
        nTamano = 1
        For i = 1 To lstLista.ListCount
            If lstLista.Selected(i - 1) = True Then
                MatLista(nTamano, 0) = Trim(Left(lstLista.List(i - 1), 3))
                nTamano = nTamano + 1
            End If
        Next
    ElseIf pnValCant = 3 Then
        ReDim MatLista(feIndCartAtrasada.Rows - 1, 3)
        For i = 1 To feIndCartAtrasada.Rows - 1
            MatLista(i - 1, 0) = Trim(feIndCartAtrasada.TextMatrix(i, 1))
            MatLista(i - 1, 1) = Trim(feIndCartAtrasada.TextMatrix(i, 2))
            MatLista(i - 1, 2) = CInt(IIf(feIndCartAtrasada.TextMatrix(i, 3) = ".", 1, 0))
        Next
    End If
    LlenaMatriz = MatLista()
End Function

Private Sub LlenaListas(ByRef Lista As ListBox, ByVal pnTipoLista As Integer)
    Dim oLista As COMDCredito.DCOMBPPR
    Dim oSubProd As COMDConstantes.DCOMConstantes
    Dim oAge As COMDConstantes.DCOMAgencias
    Dim rs As ADODB.Recordset
    Dim rsLista As ADODB.Recordset
    Dim i As Integer, J As Integer
    
    Set oLista = New COMDCredito.DCOMBPPR
    Set rs = oLista.RecuperaParametrosLista(feParametrosBPP.TextMatrix(feParametrosBPP.Row, 8), pnTipoLista)
    Set oLista = Nothing

    If pnTipoLista = 1 Then
        Set oSubProd = New COMDConstantes.DCOMConstantes
        Set rsLista = oSubProd.RecuperaConstantes(3033, 1)
        Set oSubProd = Nothing
    Else
        Set oAge = New COMDConstantes.DCOMAgencias
        Set rsLista = oAge.ObtieneAgencias()
        Set oAge = Nothing
    End If
    
    'Lista.Selected(0) = False
    
    For i = 0 To rsLista.RecordCount - 1
        Lista.AddItem rsLista!nConsValor & " " & Trim(rsLista!cConsDescripcion)
        rs.MoveFirst
        For J = 0 To rs.RecordCount - 1
            If Trim(rsLista!nConsValor) = Trim(rs!nConsValor) Then
                Lista.Selected(i) = True
            End If
            rs.MoveNext
        Next J
        rsLista.MoveNext
    Next i
End Sub
