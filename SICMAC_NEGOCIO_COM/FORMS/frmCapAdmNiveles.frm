VERSION 5.00
Begin VB.Form frmCapAdmNiveles 
   Caption         =   "Administracion de Niveles"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7545
   Icon            =   "frmCapAdmNiveles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   7545
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6300
      TabIndex        =   15
      Top             =   8085
      Width           =   1065
   End
   Begin VB.Frame Frame2 
      Caption         =   "Niveles Registrados"
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
      Height          =   3060
      Left            =   135
      TabIndex        =   16
      Top             =   4935
      Width           =   7230
      Begin SICMACT.FlexEdit grdGrupos 
         Height          =   2535
         Left            =   210
         TabIndex        =   13
         Top             =   315
         Width           =   5670
         _ExtentX        =   10001
         _ExtentY        =   4471
         Cols0           =   6
         FixedCols       =   0
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "id-Grupos-Detalles-estado-inicio-fin"
         EncabezadosAnchos=   "0-3700-1000-0-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0"
         TextArray0      =   "id"
         lbUltimaInstancia=   -1  'True
         Appearance      =   0
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5985
         TabIndex        =   14
         Top             =   840
         Width           =   1065
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5985
         TabIndex        =   12
         Top             =   315
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Agregar Niveles"
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
      Height          =   4770
      Left            =   135
      TabIndex        =   0
      Top             =   105
      Width           =   7230
      Begin VB.CommandButton btnTodos 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4935
         TabIndex        =   8
         Top             =   1155
         Width           =   960
      End
      Begin VB.CheckBox ckSubProductos 
         Caption         =   "SubProductos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3255
         TabIndex        =   7
         Top             =   1155
         Width           =   1590
      End
      Begin VB.TextBox txtDias 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2310
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1155
         Width           =   855
      End
      Begin VB.TextBox txtTasa 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         TabIndex        =   4
         Top             =   1155
         Width           =   645
      End
      Begin SICMACT.FlexEdit grdAgencia 
         Height          =   2955
         Left            =   225
         TabIndex        =   10
         Top             =   1680
         Width           =   5670
         _ExtentX        =   10001
         _ExtentY        =   5212
         Cols0           =   7
         FixedCols       =   0
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "id-Agencias-TEA Adic.-idGA-Dias-SubPr-otro"
         EncabezadosAnchos=   "0-2000-1000-0-800-800-0"
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
         ColumnasAEditar =   "X-X-2-X-4-5-X"
         ListaControles  =   "0-0-0-0-0-4-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-L-R-C-R-L-C"
         FormatosEdit    =   "0-0-2-0-3-0-0"
         CantEntero      =   5
         TextArray0      =   "id"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5985
         TabIndex        =   11
         Top             =   2205
         Width           =   1065
      End
      Begin VB.ComboBox cboGrupo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   945
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   525
         Width           =   4965
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5985
         TabIndex        =   9
         Top             =   1680
         Width           =   1065
      End
      Begin VB.Label Label3 
         Caption         =   "(Dias) > = "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1575
         TabIndex        =   5
         Top             =   1200
         Width           =   750
      End
      Begin VB.Label Label2 
         Caption         =   "Tasa:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   250
         TabIndex        =   3
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Grupo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   250
         TabIndex        =   1
         Top             =   525
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmCapAdmNiveles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************************************
'* NOMBRE         : "frmCapAdmNiveles"                                                                          *
'* DESCRIPCION    : Formulario creado para crear, editar y eliminar niveles de aprobacion de tasas              *
'*                  preferenciales, segun proyecto: "Mejora del Sistema y Automatizacion de Ahorros y Servicios"*
'* CREACION       : RIRO, 20121109 10:00 AM'                                                                    *
'****************************************************************************************************************

Option Explicit

Dim gruposUsuarios() As String  'Contiene lista de Grupo de usuarios de dominio
Dim oConecta As COMConecta.DCOMConecta

Private Sub btnTodos_Click()

    Dim i As Integer
    
    For i = 1 To grdAgencia.Rows - 1
    
        grdAgencia.TextMatrix(i, 2) = Format(txtTasa.Text, "#,##0.00")
        grdAgencia.TextMatrix(i, 4) = txtDias.Text
        grdAgencia.TextMatrix(i, 5) = ckSubProductos.value
    
    Next

End Sub

Private Sub cboGrupo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If txtTasa.Enabled Then txtTasa.SetFocus
    End If

End Sub

Private Sub Form_Load()
    
On Error GoTo error

    cargargrdGrupos
    cargarCombo
    cargargrdAgencia
    cboGrupo_Click
    txtDias.Alignment = 1
    txtTasa.MaxLength = 6
    txtDias.MaxLength = 6
    
    Exit Sub
    
error:
    MsgBox err.Description, vbCritical, "Aviso"

End Sub

Private Sub cmdCancelar_Click()

On Error GoTo error
    recargarFormulario
    Exit Sub
    
error:
    MsgBox err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cboGrupo_Click()

    If cboGrupo.Text = "" Then
        cmdGuardar.Enabled = False
        cmdCancelar.Enabled = False
        grdAgencia.Enabled = False
        txtTasa.Enabled = False
        txtDias.Enabled = False
        btnTodos.Enabled = False
        ckSubProductos.Enabled = False
        txtTasa.Text = "0.00"
        txtDias.Text = "0"
        ckSubProductos.value = 0
    Else
        cmdGuardar.Enabled = True
        cmdCancelar.Enabled = True
        grdAgencia.Enabled = True
        txtTasa.Enabled = True
        txtDias.Enabled = True
        btnTodos.Enabled = True
        ckSubProductos.Enabled = True
        
    End If

End Sub

Private Sub recargarFormulario()
    
On Error GoTo error

    cargargrdGrupos
    cargarCombo
    cargargrdAgencia
    cboGrupo_Click
    cboGrupo.SetFocus
    
    Exit Sub
    
error:
    MsgBox err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub cargarCombo()

    Dim confirm As Boolean
    Dim v As Variant
    Dim i As Integer
    
    On Error GoTo error
    
    cboGrupo.Clear
    Call cargarListaGrupos(gruposUsuarios)
    For Each v In gruposUsuarios
        confirm = False
        For i = 1 To grdGrupos.Rows - 1
           If v = grdGrupos.TextMatrix(i, 1) Then
              confirm = True
            End If
        Next
        If Not confirm Then
            cboGrupo.AddItem v
         End If
    Next
    
    Exit Sub
    
error:
    err.Raise err.Number, err.Source, err.Description
    
End Sub

Private Sub cargarListaGrupos(ByRef Lista() As String)
    
    Dim acceso As COMDPersona.UCOMAcceso
    
    On Error GoTo error
    
    Set acceso = New COMDPersona.UCOMAcceso
    
    Lista = Split(acceso.listGrupos(gsDominio), ",")
    
    Exit Sub
    
error:
    err.Raise err.Number, err.Source, err.Description
    
End Sub

Private Sub grdAgencia_OnCellChange(pnRow As Long, pnCol As Long)

 If grdAgencia.col = 2 Then
    
        If Not IsNumeric(grdAgencia.TextMatrix(grdAgencia.row, 2)) Then
        
            MsgBox "Valor de la celda no es numerico", vbExclamation, "Aviso"
            grdAgencia.TextMatrix(grdAgencia.row, 2) = "0.00"
            
        Else
            
            If grdAgencia.TextMatrix(grdAgencia.row, 2) < 0 Then
            
             MsgBox "Valor de la celda es menor que 0.00", vbExclamation, "Aviso"
             grdAgencia.TextMatrix(grdAgencia.row, 2) = "0.00"
            
            ElseIf grdAgencia.TextMatrix(grdAgencia.row, 2) > 100 Then
            
             MsgBox "Valor de la celda no puede ser mayor que 100.00", vbExclamation, "Aviso"
             grdAgencia.TextMatrix(grdAgencia.row, 2) = "0.00"
            
            End If
                        
        End If
    
    End If

End Sub

Private Sub grdGrupos_Click()

    Dim row As Integer
    Dim col As Integer
    Dim Index As Integer
    Dim Sql As String
    Dim oFrmDialog As frmCapAdmNivelesDialog
    Dim oCaptaServicios As COMDCaptaServicios.DCOMCaptaServicios
    Dim rsAge As ADODB.Recordset
    
    On Error GoTo error
    
    If Not grdGrupos.TextMatrix(1, 1) = "" Then
        row = grdGrupos.row
        col = grdGrupos.col
        If row >= 1 And col = 2 Then
            Set oCaptaServicios = New COMDCaptaServicios.DCOMCaptaServicios
            Set oFrmDialog = New frmCapAdmNivelesDialog
            Index = grdGrupos.TextMatrix(row, 0)
            Set rsAge = oCaptaServicios.getAgenciaTea(Index)
            oFrmDialog.cargarRs rsAge
            oFrmDialog.Frame1.Caption = grdGrupos.TextMatrix(row, 1)
            oFrmDialog.Show 1
            Set oCaptaServicios = Nothing
            Set oFrmDialog = Nothing
            Set rsAge = Nothing
        End If
        
    End If
    
    Exit Sub
    
error:
    MsgBox err.Description, vbCritical, "Aviso"

End Sub

Private Sub CmdQuitar_Click()

    Dim oCaptaServicios As COMDCaptaServicios.DCOMCaptaServicios
    Dim Index As Integer
    Dim sMensaje As String
    
    sMensaje = "Desea eliminar el grupo: " & grdGrupos.TextMatrix(grdGrupos.row, 1) & " de la lista de niveles aprobados?"
    
    If MsgBox(sMensaje, vbYesNo + vbInformation, "Aviso") = vbNo Then
        Exit Sub
    End If
    
On Error GoTo error
    
    Set oCaptaServicios = New COMDCaptaServicios.DCOMCaptaServicios
    Index = grdGrupos.TextMatrix(grdGrupos.row, 0)
    Call oCaptaServicios.quitarGrupoAgencia(Index, Format(gdFecSis, "yyyyMMdd"))
    
    Set oCaptaServicios = Nothing
    recargarFormulario
    MsgBox "Se procedio a retirar el grupo seleccionado de la lista de niveles aprobados", vbInformation, "Aviso"
    cboGrupo.SetFocus
    
    Exit Sub
    
error:
    MsgBox err.Description, vbCritical, "Aviso"

End Sub

Private Sub CmdEditar_Click()
    
    Dim Index, i As Integer
    Dim oCaptaServicios As COMDCaptaServicios.DCOMCaptaServicios
    
On Error GoTo error
    
    cmdQuitar.Enabled = False
    cmdEditar.Enabled = False
    cboGrupo.Clear
    cboGrupo.AddItem grdGrupos.TextMatrix(grdGrupos.row, 1)
    cmdGuardar.Caption = "Editar"
    cboGrupo.ListIndex = 0
    grdGrupos.Enabled = False
    Index = grdGrupos.TextMatrix(grdGrupos.row, 0)
    Set oCaptaServicios = New COMDCaptaServicios.DCOMCaptaServicios
    grdAgencia.Clear
    grdAgencia.rsFlex = oCaptaServicios.getAgenciaTea(Index)
    For i = 1 To grdAgencia.Rows - 1
        grdAgencia.TextMatrix(i, 2) = Format$(grdAgencia.TextMatrix(i, 2), "#,##0.00")
        grdAgencia.TextMatrix(i, 5) = IIf(grdAgencia.TextMatrix(i, 5) = ".", 1, "")
    Next
    grdAgencia.SetFocus
    grdAgencia.row = 1
    grdAgencia.col = 2
    
    Exit Sub
            
error:
    MsgBox err.Description, vbCritical, "Aviso"
            
End Sub

Private Sub cmdGuardar_Click()

    Dim row As Integer
    Dim idGrupo As Integer
    Dim i As Integer
    Dim oCaptaServicios As COMDCaptaServicios.DCOMCaptaServicios
    
    On Error GoTo error
    
    row = grdAgencia.Rows - 1
    
    If cmdGuardar.Caption = "Guardar" Then
    
        If MsgBox("Desea registrar un nuevo nivel de aprobación?", vbYesNo + vbInformation, "Aviso") = vbYes Then
            idGrupo = agregarGrupo(cboGrupo.Text)
            For i = 1 To row
                Call agregarGrupoAgencia(grdAgencia.TextMatrix(i, 0), idGrupo, CDbl(grdAgencia.TextMatrix(i, 2)), _
                 val(grdAgencia.TextMatrix(i, 4)), IIf(grdAgencia.TextMatrix(i, 5) = ".", 1, 0))
            Next
            recargarFormulario
            MsgBox "El nuevo nivel de aprobación fue registrado correctamente", vbInformation, "Aviso"
            cboGrupo.SetFocus
        End If
        
    ElseIf cmdGuardar.Caption = "Editar" Then
    
        If MsgBox("Desea modificar las tasas del nivel de aprobacion: " & cboGrupo.Text & " ?", vbYesNo + vbInformation, "Aviso") = vbYes Then
            
            Set oCaptaServicios = New COMDCaptaServicios.DCOMCaptaServicios
            For i = 1 To grdAgencia.Rows - 1
                oCaptaServicios.editarGrupoAgencia Format$(ConvierteTEAaTNA(grdAgencia.TextMatrix(i, 2)), "#0.000000"), _
                grdAgencia.TextMatrix(i, 3), val(grdAgencia.TextMatrix(i, 4)), IIf(grdAgencia.TextMatrix(i, 5) = ".", 1, 0), _
                gsCodUser, Format(gdFecSis, "yyyyMMdd")
            Next
            Set oCaptaServicios = Nothing
            recargarFormulario
            MsgBox "Las tasas se editaron correctamente", vbInformation, "Aviso"
            
        End If
        
    End If
    
    Exit Sub
    
error:
    MsgBox err.Description, vbCritical, "Aviso"
    
End Sub

Private Function agregarGrupo(ByVal grupoNuevo As String) As Integer
    
    Dim oCaptaServicios As COMDCaptaServicios.DCOMCaptaServicios
    Dim rsGrupo As ADODB.Recordset
    
    On Error GoTo error
    
    Set oCaptaServicios = New COMDCaptaServicios.DCOMCaptaServicios
    Set rsGrupo = oCaptaServicios.insertarGrupo(grupoNuevo, Format(gdFecSis, "yyyyMMdd"))
    
    If Not rsGrupo.EOF Then
        agregarGrupo = rsGrupo!valor
    End If
    
    Set oCaptaServicios = Nothing
    Set rsGrupo = Nothing
    
    Exit Function
    
error:
    err.Raise err.Number, err.Source, err.Description
    
End Function

Private Sub agregarGrupoAgencia(ByVal idAgencia As String, ByVal idGrupo As Integer, ByVal TEAAdicional As Double, _
nDias As Double, nSubProducto As Integer)

    Dim oCaptaServicios As COMDCaptaServicios.DCOMCaptaServicios
    
On Error GoTo error
    
    Set oCaptaServicios = New COMDCaptaServicios.DCOMCaptaServicios
    TEAAdicional = ConvierteTEAaTNA(TEAAdicional)
    TEAAdicional = Format$(TEAAdicional, "#0.000000")
    oCaptaServicios.insertarGrupoAgencia idAgencia, idGrupo, TEAAdicional, nDias, nSubProducto, gsCodUser, Format(gdFecSis, "yyyyMMdd")
    Set oCaptaServicios = Nothing
    
    Exit Sub
    
error:
    err.Raise err.Number, err.Source, err.Description
 
End Sub

Private Sub cargargrdGrupos()
    
    Dim oCaptaServicios As COMDCaptaServicios.DCOMCaptaServicios
    
    On Error GoTo error
    
    Set oCaptaServicios = New COMDCaptaServicios.DCOMCaptaServicios
    cmdGuardar.Enabled = True
    cmdEditar.Enabled = True
    cmdQuitar.Enabled = True
    grdGrupos.Enabled = True
    grdGrupos.Clear
    
    grdGrupos.rsFlex = oCaptaServicios.getGruposDominio
    Set oCaptaServicios = Nothing
    If grdGrupos.TextMatrix(1, 1) = "" Then
        cmdEditar.Enabled = False
        cmdQuitar.Enabled = False
        grdGrupos.Enabled = False
    End If
    
    Exit Sub
    
error:
    err.Raise err.Number, err.Source, err.Description
    
End Sub

Private Sub cargargrdAgencia()
    
    Dim i As Integer
    Dim oCaptaServicios As COMDCaptaServicios.DCOMCaptaServicios
    
On Error GoTo error
    
    Set oCaptaServicios = New COMDCaptaServicios.DCOMCaptaServicios
    grdAgencia.rsFlex = oCaptaServicios.getAgencia
    Set oCaptaServicios = Nothing
    For i = 1 To grdAgencia.Rows - 1
        
        grdAgencia.TextMatrix(i, 2) = "0"
        grdAgencia.TextMatrix(i, 4) = "0"
        grdAgencia.TextMatrix(i, 2) = Format$(grdAgencia.TextMatrix(i, 2), "#,##0.00")
        grdAgencia.TextMatrix(i, 4) = Format$(grdAgencia.TextMatrix(i, 4), "#0")
        
    Next
    cmdGuardar.Caption = "Guardar"
    
    Exit Sub
    
error:
    err.Raise err.Number, err.Source, err.Description
    
End Sub

Private Sub txtDias_GotFocus()
    txtDias.SelStart = 0
    txtDias.SelLength = Len(txtDias.Text)
    txtDias.SetFocus
End Sub

Private Sub txtDias_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        btnTodos.SetFocus
        
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
        
    End If

End Sub

Private Sub txtTasa_Change()
    
    If Not IsNumeric(txtTasa.Text) Then
        MsgBox "Debes ingresar valores numericos", vbInformation, "Aviso"
        txtTasa.Text = "0.00"
        Exit Sub
    End If
    
    If val(Trim(txtTasa.Text)) > 100 Then
        txtTasa.Text = "100.00"
    End If
    
End Sub

Private Sub txtTasa_GotFocus()
    txtTasa.SelStart = 0
    txtTasa.SelLength = Len(txtTasa.Text)
    txtTasa.SetFocus
End Sub

Private Sub txtTasa_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtDias.SetFocus
    
    Else
        KeyAscii = NumerosDecimales(txtTasa, KeyAscii)
            
    End If

End Sub

Private Sub txtTasa_LostFocus()
    txtTasa.Text = Format(txtTasa.Text, "#,##00.00")
End Sub
