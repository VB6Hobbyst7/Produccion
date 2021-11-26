VERSION 5.00
Begin VB.Form frmCapParametros 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   Icon            =   "frmCapParametros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   " Categoria "
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
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6735
      Begin VB.ComboBox cboCategoria 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   4740
      TabIndex        =   2
      Top             =   5325
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   3660
      TabIndex        =   1
      Top             =   5325
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5820
      TabIndex        =   3
      Top             =   5325
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2580
      TabIndex        =   5
      Top             =   5325
      Width           =   990
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5325
      Width           =   975
   End
   Begin VB.Frame fraParam 
      Height          =   4335
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   6735
      Begin SICMACT.FlexEdit grdParam 
         Height          =   4005
         Left            =   105
         TabIndex        =   0
         Top             =   210
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   7064
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Constante-Descripción-Valor-Tag"
         EncabezadosAnchos=   "350-0-4500-1200-0"
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
         ColumnasAEditar =   "X-X-X-3-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-L-R-C"
         FormatosEdit    =   "0-0-0-2-0"
         CantDecimales   =   3
         AvanceCeldas    =   1
         TextArray0      =   "#"
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
      End
   End
End
Attribute VB_Name = "frmCapParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'By capi 21012009
Dim objPista As COMManejador.Pista
'
'JUEZ 20141115 ******************
Dim fbConsulta As Boolean
Dim fbPermisoLavado As Boolean
Dim fbPermisoCred As Boolean
Dim fbPermisoAhorros As Boolean
'END JUEZ ***********************
Dim fbPermisoComisiones As Boolean 'JUEZ 20151229

Public Sub Inicia(Optional bConsulta As Boolean = False)
fbConsulta = bConsulta 'JUEZ 20141115
If bConsulta Then
    cmdCancelar.Visible = False
    cmdGrabar.Visible = False
    cmdEditar.Visible = False
    Me.Caption = "Captaciones - Parámetros - Consulta"
    grdParam.lbEditarFlex = False
Else
    cmdImprimir.Visible = False
    cmdCancelar.Visible = True
    cmdGrabar.Visible = True
    cmdEditar.Visible = True
    cmdCancelar.Enabled = False
    cmdGrabar.Enabled = False
    Me.Caption = "Captaciones - Parámetros - Mantenimiento"
End If
Me.Show 1
End Sub

Private Sub cboCategoria_Click()
CargarParametros CInt(Right(Trim(cboCategoria), 2))
End Sub

Private Sub cmdCancelar_Click()
cmdEditar.Enabled = True
cmdCancelar.Enabled = False
cmdGrabar.Enabled = False
cmdSalir.Enabled = True
grdParam.SetFocus
End Sub

Private Sub CmdEditar_Click()
cmdEditar.Enabled = False
cmdCancelar.Enabled = True
cmdGrabar.Enabled = True
cmdSalir.Enabled = False
grdParam.lbEditarFlex = True
End Sub

Private Sub cmdGrabar_Click()
If MsgBox("¿Desea grabar la información actualizada?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
    Exit Sub
End If
Dim clsparam As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim i As Integer
Set clsparam = New COMNCaptaGenerales.NCOMCaptaDefinicion
'By Capi 01042008
Dim lnPorRetCTS  As Integer
Dim lnPorRetAdiCTS As Integer

For i = 1 To grdParam.Rows - 1
    If grdParam.TextMatrix(i, 1) = 2021 Then
        lnPorRetCTS = grdParam.TextMatrix(i, 3)
    End If
    If grdParam.TextMatrix(i, 1) = 2099 Then
        lnPorRetAdiCTS = grdParam.TextMatrix(i, 3)
    End If
Next i
'
For i = 1 To grdParam.Rows - 1
    If grdParam.TextMatrix(i, 4) = "M" Then
        'By Capi 01042008 para que controle parametros CTS
        If grdParam.TextMatrix(i, 1) = 2021 Or grdParam.TextMatrix(i, 1) = 2099 Then
            If lnPorRetCTS + lnPorRetAdiCTS > 100 Then
                MsgBox "Parametros Porcentaje Retiro CTS  y Porcentaje Adicional Retiro CTS no son Validos, superan 100%, verifique parametros"
                Exit Sub
            End If
        End If
        
        '
        clsparam.ActualizaParametros grdParam.TextMatrix(i, 1), grdParam.TextMatrix(i, 2), CDbl(grdParam.TextMatrix(i, 3))
        'By Capi 21012009
        objPista.InsertarPista gsOpeCod, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gModificar, "Tabla Parametro : " & grdParam.TextMatrix(i, 1)
        '
    End If
Next i
Set clsparam = Nothing
cmdCancelar.Enabled = False
cmdGrabar.Enabled = False
cmdSalir.Enabled = True
cmdEditar.Enabled = True
grdParam.SetFocus
End Sub

Private Sub cmdImprimir_Click()
Dim sCad As String
Dim Prev As previo.clsprevio
Dim rs As New ADODB.Recordset
Dim LsCapImp As COMNCaptaGenerales.NCOMCaptaImpresion
Dim i As Integer

With rs
   'Crear RecordSet
    .Fields.Append "cDescripcion", adVarChar, 250
    .Fields.Append "cValor", adVarChar, 250
    .Open
    'Llenar Recordset
    For i = 1 To grdParam.Rows - 1
        .AddNew
        .Fields("cDescripcion") = grdParam.TextMatrix(i, 2)
        .Fields("cValor") = grdParam.TextMatrix(i, 3)
    Next i
End With

Set LsCapImp = New COMNCaptaGenerales.NCOMCaptaImpresion
    sCad = LsCapImp.ImprimirParametros(rs, gsNomAge, gdFecSis)
Set LsCapImp = Nothing
Set Prev = New previo.clsprevio
Prev.Show sCad, "Parámetros Captaciones", True, , gImpresora
Set Prev = Nothing
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
'JUEZ  ******************************************
'Dim clsparam  As COMNCaptaGenerales.NCOMCaptaDefinicion
'Dim rsPar As ADODB.Recordset
'Set clsparam = New COMNCaptaGenerales.NCOMCaptaDefinicion
'Set rsPar = New ADODB.Recordset
Dim oGen As COMDConstSistema.DCOMGeneral
Me.Icon = LoadPicture(App.path & gsRutaIcono)

fbPermisoLavado = False
fbPermisoCred = False
Set oGen = New COMDConstSistema.DCOMGeneral
fbPermisoLavado = oGen.VerificaExistePermisoCargo(gsCodCargo, gMantParamLavado)
fbPermisoCred = oGen.VerificaExistePermisoCargo(gsCodCargo, gMantParamCreditos)
fbPermisoAhorros = oGen.VerificaExistePermisoCargo(gsCodCargo, gMantParamAhorros)
fbPermisoComisiones = oGen.VerificaExistePermisoCargo(gsCodCargo, gMantParamComisiones) 'JUEZ 20151229

CargarCategorias
'rsPar.CursorLocation = adUseClient
'Set rsPar = clsparam.GetParametros
'grdParam.rsFlex = rsPar
'Set rsPar = Nothing
'Set clsparam = Nothing
If Trim(cboCategoria.Text) <> "" Then CargarParametros (CInt(Right(Trim(cboCategoria.Text), 2)))
'END JUEZ ***********************************************
'By Capi 20012009
Set objPista = New COMManejador.Pista
gsOpeCod = gCapMantParametros
'End By
End Sub

Private Sub grdParam_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmdGrabar.Visible Then
        CmdEditar_Click
    Else
        cmdImprimir.SetFocus
    End If
    'grdParam.TextMatrix(grdParam.Row, 4) = "M"
    If grdParam.Col = 3 Then
        If grdParam.TextMatrix(grdParam.row, 3) < 0 Then
            MsgBox "Ingrese solo valores positivos", vbInformation, "Aviso"
        End If
    End If
End If
End Sub

Private Sub grdParam_OnCellChange(pnRow As Long, pnCol As Long)
    grdParam.TextMatrix(pnRow, 4) = "M"
    If grdParam.Col = 3 Then
        If grdParam.TextMatrix(grdParam.row, 3) < 0 Then
            MsgBox "Ingrese solo valores positivos", vbInformation, "Aviso"
        End If
    End If
End Sub

Private Sub grdParam_OnRowChange(pnRow As Long, pnCol As Long)
''grdParam.TextMatrix(pnRow, 4) = "M"
'If grdParam.Col = 3 Then
'    If grdParam.TextMatrix(grdParam.Row, 3) < 0 Then
'        MsgBox "Ingrese solo valores positivos", vbInformation, "Aviso"
'    End If
'End If
End Sub

Private Sub grdParam_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    If grdParam.Col = 3 Then
    If grdParam.TextMatrix(grdParam.row, 3) < 0 Then
        MsgBox "Ingrese solo valores positivos", vbInformation, "Aviso"
    End If
End If

End Sub
'JUEZ 20141115 *******************************************
Private Sub CargarCategorias()
Dim clsGen As COMDConstSistema.DCOMGeneral
Dim rsConst As ADODB.Recordset
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsConst = clsGen.GetConstante(2041)
    Set clsGen = Nothing
    Do While Not rsConst.EOF
        If Not fbConsulta Then
            If rsConst("nConsValor") = "6" And fbPermisoLavado Then
                cboCategoria.AddItem rsConst("cDescripcion") & Space(500) & rsConst("nConsValor")
                cboCategoria.ListIndex = IndiceListaCombo(cboCategoria, rsConst("nConsValor"))
            End If
            If fbPermisoCred And rsConst("nConsValor") = "7" Then
                cboCategoria.AddItem rsConst("cDescripcion") & Space(500) & rsConst("nConsValor")
                cboCategoria.ListIndex = IndiceListaCombo(cboCategoria, rsConst("nConsValor"))
            End If
            'If fbPermisoAhorros And rsConst("nConsValor") <> "6" And rsConst("nConsValor") <> "7" Then
            If fbPermisoAhorros And rsConst("nConsValor") <> "6" And rsConst("nConsValor") <> "7" And rsConst("nConsValor") <> "14" Then 'JUEZ 20151229
                cboCategoria.AddItem rsConst("cDescripcion") & Space(500) & rsConst("nConsValor")
                cboCategoria.ListIndex = 0
            End If
            'JUEZ 20151229 ***************************************************************
            If fbPermisoComisiones And rsConst("nConsValor") = "14" Then
                cboCategoria.AddItem rsConst("cDescripcion") & Space(500) & rsConst("nConsValor")
                cboCategoria.ListIndex = IndiceListaCombo(cboCategoria, rsConst("nConsValor"))
            End If
            'END JUEZ ********************************************************************
        Else
            cboCategoria.AddItem rsConst("cDescripcion") & Space(500) & rsConst("nConsValor")
            cboCategoria.ListIndex = 0
        End If
        rsConst.MoveNext
    Loop
End Sub
Private Sub CargarParametros(Optional ByVal pnCategoria As Integer = 0)
Dim clsparam  As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim rsPar As ADODB.Recordset
    Set clsparam = New COMNCaptaGenerales.NCOMCaptaDefinicion
    Set rsPar = New ADODB.Recordset
    rsPar.CursorLocation = adUseClient
    Set rsPar = clsparam.GetParametros(pnCategoria)
    grdParam.rsFlex = rsPar
    Set rsPar = Nothing
    Set clsparam = Nothing
End Sub
'END JUEZ ************************************************
