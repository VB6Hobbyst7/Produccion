VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmMantCodigoPostal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento Codigo Postal"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8100
   Icon            =   "frmMantCodigoPostal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   1515
      Left            =   60
      TabIndex        =   8
      Top             =   4140
      Width           =   7995
      Begin MSDataListLib.DataCombo cmbDistrito 
         Height          =   315
         Left            =   5010
         TabIndex        =   11
         Top             =   1110
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         MaxLength       =   25
         TabIndex        =   10
         Top             =   435
         Width           =   3915
      End
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   105
         TabIndex        =   9
         Top             =   435
         Width           =   900
      End
      Begin MSDataListLib.DataCombo cmbDepartamento 
         Height          =   315
         Left            =   90
         TabIndex        =   15
         Top             =   1110
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbProvincia 
         Height          =   315
         Left            =   2340
         TabIndex        =   17
         Top             =   1110
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label5 
         Caption         =   "Provincia"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   2370
         TabIndex        =   18
         Top             =   900
         Width           =   840
      End
      Begin VB.Label Label4 
         Caption         =   "Departamento"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   915
         Width           =   1260
      End
      Begin VB.Label Label3 
         Caption         =   "Distrito"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   5025
         TabIndex        =   14
         Top             =   915
         Width           =   840
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   1170
         TabIndex        =   13
         Top             =   225
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   225
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4155
      Left            =   60
      TabIndex        =   7
      Top             =   0
      Width           =   7995
      Begin SICMACT.FlexEdit feCodigoPostal 
         Height          =   3900
         Left            =   75
         TabIndex        =   19
         Top             =   180
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   6879
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Codigo-Descripcion-Zona-CodZona"
         EncabezadosAnchos=   "0-800-3000-3900-0"
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   60
      TabIndex        =   0
      Top             =   5640
      Width           =   7995
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   390
         Left            =   150
         TabIndex        =   1
         Top             =   225
         Width           =   1185
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
         Height          =   390
         Left            =   1365
         TabIndex        =   2
         Top             =   225
         Width           =   1185
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   390
         Left            =   2580
         TabIndex        =   3
         Top             =   225
         Width           =   1185
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   390
         Left            =   5400
         TabIndex        =   4
         Top             =   225
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   390
         Left            =   6615
         TabIndex        =   5
         Top             =   225
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   6615
         TabIndex        =   6
         Top             =   240
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmMantCodigoPostal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnNuevoEditar As Integer

Private Sub cmbDepartamento_Click(Area As Integer)

If cmbDepartamento.BoundText <> "" Then
    CargaCombo cmbProvincia, "2", Mid(cmbDepartamento.BoundText, 2, 2)
End If
End Sub

Private Sub cmbDepartamento_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    If cmbDepartamento.BoundText <> "" Then
        CargaCombo cmbProvincia, "2", Mid(cmbDepartamento.BoundText, 2, 2)
        cmbProvincia.SetFocus
    End If
End If

End Sub

Private Sub cmbDistrito_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdAceptar.SetFocus
    End If
End Sub

Private Sub cmbProvincia_Click(Area As Integer)
    If cmbProvincia.BoundText <> "" Then
        CargaCombo cmbDistrito, "3", Mid(cmbDepartamento.BoundText, 2, 2), Mid(cmbProvincia.BoundText, 4, 2)
    End If
End Sub

Private Sub cmbProvincia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmbProvincia.BoundText <> "" Then
        CargaCombo cmbDistrito, "3", Mid(cmbDepartamento.BoundText, 2, 2), Mid(cmbProvincia.BoundText, 4, 2)
        cmbDistrito.SetFocus
    End If
End If
End Sub

Private Sub CmdAceptar_Click()
Dim oGen As COMDPersona.DCOMPersGeneral

If lnNuevoEditar = 1 Then   'Nuevo
    If ValidaGrabar Then
    
        Set oGen = New COMDPersona.DCOMPersGeneral
        oGen.dInsertCodigoPostal TxtCodigo.Text, TxtDescripcion.Text, cmbDistrito.BoundText
        
        feCodigoPostal.AdicionaFila
        feCodigoPostal.TextMatrix(feCodigoPostal.Rows - 1, 1) = TxtCodigo
        feCodigoPostal.TextMatrix(feCodigoPostal.Rows - 1, 2) = TxtDescripcion
        feCodigoPostal.TextMatrix(feCodigoPostal.Rows - 1, 3) = cmbDistrito.Text
        feCodigoPostal.TextMatrix(feCodigoPostal.Rows - 1, 4) = cmbDistrito.BoundText
         
        Set oGen = Nothing
    Else
        Exit Sub
    End If
ElseIf lnNuevoEditar = 2 Then           'Editar
    If ValidaGrabar Then
        
        Set oGen = New COMDPersona.DCOMPersGeneral
        oGen.dUpdateCodigoPostal TxtCodigo, TxtDescripcion, cmbDistrito.BoundText
        
        feCodigoPostal.TextMatrix(feCodigoPostal.Row, 1) = TxtCodigo
        feCodigoPostal.TextMatrix(feCodigoPostal.Row, 2) = TxtDescripcion
        feCodigoPostal.TextMatrix(feCodigoPostal.Row, 3) = cmbDistrito.Text
        feCodigoPostal.TextMatrix(feCodigoPostal.Row, 4) = cmbDistrito.BoundText
        
        Set oGen = Nothing
    Else
        Exit Sub
    End If
End If

Limpia
HabilitaBotones True, True, True, False, False, True
Frame2.Enabled = False

End Sub

Private Sub cmdCancelar_Click()
    Limpia
    HabilitaBotones True, True, True, False, False, True
    Frame2.Enabled = False
End Sub

Private Sub CmdEditar_Click()

    lnNuevoEditar = 2
    MuestraDatos
    Frame2.Enabled = True
    HabilitaBotones False, False, False, True, True, False
    TxtDescripcion.SetFocus
    
End Sub

Private Sub cmdeliminar_Click()
Dim oGen As COMDPersona.DCOMPersGeneral

If MsgBox("Desea Eliminar el Codigo Postal seleccionado?", vbYesNo, "Aviso") = vbYes Then

    Set oGen = New COMDPersona.DCOMPersGeneral
    
    oGen.dDeleteCodigoPostal feCodigoPostal.TextMatrix(feCodigoPostal.Row, 1)
    feCodigoPostal.EliminaFila feCodigoPostal.Row
    
    Set oGen = Nothing
    
End If

End Sub

Private Sub cmdNuevo_Click()
Dim oGen As COMDPersona.DCOMPersGeneral

Limpia
lnNuevoEditar = 1
Frame2.Enabled = True
HabilitaBotones False, False, False, True, True, False

Set oGen = New COMDPersona.DCOMPersGeneral
    TxtCodigo.Text = oGen.dGeneraCodigoPostal

Set oGen = Nothing
TxtDescripcion.SetFocus

End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub feCodigoPostal_Click()
    MuestraDatos
End Sub

Private Sub Form_Load()
Dim oGen As COMDPersona.DCOMPersGeneral
Dim rs As Recordset
Me.Icon = LoadPicture(App.path & gsRutaIcono)
lnNuevoEditar = 0
Set oGen = New COMDPersona.DCOMPersGeneral
Set rs = oGen.dObtieneCodigosPostal

If Not rs.EOF And Not rs.BOF Then
    Do While Not rs.EOF
        feCodigoPostal.AdicionaFila
        feCodigoPostal.TextMatrix(feCodigoPostal.Rows - 1, 1) = rs!nCodPostal
        feCodigoPostal.TextMatrix(feCodigoPostal.Rows - 1, 2) = rs!cDesCodPostal
        feCodigoPostal.TextMatrix(feCodigoPostal.Rows - 1, 3) = rs!cUbiGeoDescripcion
        feCodigoPostal.TextMatrix(feCodigoPostal.Rows - 1, 4) = rs!cCodZon
        rs.MoveNext
    Loop
End If

Set rs = Nothing
Set oGen = Nothing

CargaCombo cmbDepartamento, "1"

End Sub

Private Sub CargaCombo(Combo As DataCombo, ByVal psTipoZon As String, Optional ByVal psDep As String = "@", _
                        Optional ByVal psProv As String = "@")
                        
Dim oGen As COMDPersona.DCOMPersGeneral
Dim rs As Recordset

Set oGen = New COMDPersona.DCOMPersGeneral

    Set rs = oGen.dObtieneZona(psTipoZon, psDep, psProv)
    Set Combo.RowSource = rs
    Combo.ListField = "cUbiGeoDescripcion"
    Combo.BoundColumn = "cUbiGeoCod"

    Set rs = Nothing
    Set oGen = Nothing

End Sub

Private Function ValidaGrabar() As Boolean
ValidaGrabar = True

If TxtDescripcion = "" Then
    MsgBox "Debe Ingresar la Descripción del Código Postal", vbInformation, "Aviso"
    ValidaGrabar = False
    Exit Function
End If

If cmbDistrito.BoundText = "" Then
    MsgBox "Debe seleccionar un distrito para el Código Postal", vbInformation, "Aviso"
    ValidaGrabar = False
    Exit Function
End If

End Function

Private Sub Limpia()

    TxtCodigo.Text = ""
    TxtDescripcion.Text = ""
    cmbDepartamento.BoundText = ""
    cmbProvincia.BoundText = ""
    cmbDistrito.BoundText = ""
    
End Sub

Private Sub HabilitaBotones(ByVal pbNuevo As Boolean, ByVal pbEditar As Boolean, ByVal pbEliminar As Boolean, _
                            ByVal pbAceptar As Boolean, ByVal pbCancelar As Boolean, ByVal pbSalir As Boolean)

    CmdNuevo.Visible = pbNuevo
    CmdEditar.Visible = pbEditar
    CmdEliminar.Visible = pbEliminar
    CmdAceptar.Visible = pbAceptar
    CmdCancelar.Visible = pbCancelar
    CmdSalir.Visible = pbSalir
    
End Sub

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        cmbDepartamento.SetFocus
    End If
    
End Sub

Private Sub MuestraDatos()
Dim oGen As COMDPersona.DCOMPersGeneral
Dim rs As Recordset

    If feCodigoPostal.TextMatrix(feCodigoPostal.Row, 1) <> "" Then
        
        Set oGen = New COMDPersona.DCOMPersGeneral
        
        TxtCodigo = feCodigoPostal.TextMatrix(feCodigoPostal.Row, 1)
        TxtDescripcion = feCodigoPostal.TextMatrix(feCodigoPostal.Row, 2)
        Set rs = oGen.dObtieneZona(1, Mid(feCodigoPostal.TextMatrix(feCodigoPostal.Row, 4), 2, 2))
        cmbDepartamento.BoundText = rs!cUbiGeoCod
        Set rs = Nothing
        Call cmbDepartamento_Click(1)
        Set rs = oGen.dObtieneZona(2, Mid(feCodigoPostal.TextMatrix(feCodigoPostal.Row, 4), 2, 2), Mid(feCodigoPostal.TextMatrix(feCodigoPostal.Row, 4), 4, 2))
        cmbProvincia.BoundText = rs!cUbiGeoCod
        Set rs = Nothing
        Call cmbProvincia_Click(1)
        cmbDistrito.BoundText = feCodigoPostal.TextMatrix(feCodigoPostal.Row, 4)
            
        Set oGen = Nothing
        
    End If

End Sub
