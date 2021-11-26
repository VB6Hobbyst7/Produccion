VERSION 5.00
Begin VB.Form frmProveedorRegSistemaPensionLista 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema Pensión de Proveedor"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11145
   Icon            =   "frmProveedorRegSistemaPensionLista.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   11145
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   320
      Left            =   2105
      TabIndex        =   4
      Top             =   4050
      Width           =   975
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   320
      Left            =   120
      TabIndex        =   2
      Top             =   4050
      Width           =   975
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   320
      Left            =   1110
      TabIndex        =   1
      Top             =   4050
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   320
      Left            =   10080
      TabIndex        =   0
      Top             =   4050
      Width           =   975
   End
   Begin Sicmact.FlexEdit fg 
      Height          =   3885
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   6853
      Cols0           =   7
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-cProveedor-Proveedor-Sistema Pensión-AFP: CUSP-AFP: Entidad-AFP: Tipo Comisión"
      EncabezadosAnchos=   "400-0-3000-2000-1500-2000-1800"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-L-C-L-L"
      FormatosEdit    =   "0-0-0-0-0-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      lbPuntero       =   -1  'True
      lbOrdenaCol     =   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmProveedorRegSistemaPensionLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************
'** Nombre : frmProveedorRegSistemaPensionLista
'** Descripción : Formulario para listar los proveedores que están aplicando retención
'** Creación : EJVG, 20140801 12:00:00 PM
'*************************************************************************************
Option Explicit

Private Sub cmdBuscar_Click()
    Dim lsPersCod As String
    Dim lsBusca As String
    Dim fila As Integer
    
    On Error GoTo ErrBuscar
    If fg.TextMatrix(1, 1) = "" Then Exit Sub
    
    lsBusca = Trim(InputBox("Ingrese el Nombre de la Persona a buscar", "Buscar.."))
    If lsBusca = "" Then Exit Sub
    
    For fila = 1 To fg.Rows - 1
        If UCase(lsBusca) = UCase(Left(fg.TextMatrix(fila, 2), Len(lsBusca))) Then
            fg.TopRow = fila
            fg.row = fila
            Exit For
        End If
    Next
    Exit Sub
ErrBuscar:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdEditar_Click()
    Dim frm As frmProveedorRegSistemaPension
    Dim lsPersCod As String
    
    On Error GoTo ErrNuevo
    If fg.TextMatrix(1, 1) = "" Then
        MsgBox "Ud. debe seleccionar primero al Proveedor", vbInformation, "Aviso"
        Exit Sub
    End If
    lsPersCod = fg.TextMatrix(fg.row, 1)
    Set frm = New frmProveedorRegSistemaPension
    frm.Editar (lsPersCod)
    If frm.bOK Then
        cargar_datos
    End If
    Set frm = Nothing
    Exit Sub
ErrNuevo:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdNuevo_Click()
    Dim frm As frmProveedorRegSistemaPension
    Dim frmBusca As New frmBuscaPersona
    Dim oPersona As UPersona
    Dim oNProv As NProveedorSistPens
    Dim oNCons As NConstSistemas
    Dim lnRMV As Currency

    On Error GoTo ErrNuevo

    Set oPersona = frmBusca.Inicio
    If Not oPersona Is Nothing Then
        If oPersona.sPersCod <> "" Then
            Set oNProv = New NProveedorSistPens
            Set oNCons = New NConstSistemas
            lnRMV = oNCons.LeeConstSistema(480) 'Ingresamos el monto minimo que se encuentra en la validación
            If oNProv.AplicaRetencionSistemaPension(oPersona.sPersCod, gdFecSis, lnRMV) Then
                If Not oNProv.ExisteDatosSistemaPension(oPersona.sPersCod) Then
                    Set frm = New frmProveedorRegSistemaPension
                    frm.Registrar (oPersona.sPersCod)
                    If frm.bOK Then
                        cargar_datos
                    End If
                Else
                    MsgBox "La persona seleccionada ya cuenta con Datos de Sistema de Pensión." & Chr(10) & "Ud. debe editar a esta persona", vbInformation, "Aviso"
                End If
            Else
                MsgBox "La persona seleccionada no cumple los requisitos para ingresar los datos de Sistema de Pensión", vbInformation, "Aviso"
            End If
        End If
    End If

    Set oNCons = Nothing
    Set oNProv = Nothing
    Set oPersona = Nothing
    Set frmBusca = Nothing
    Set frm = Nothing
    Exit Sub
ErrNuevo:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    cargar_datos
End Sub
Private Function cargar_datos() As Boolean
    Dim obj As New NProveedorSistPens
    Dim rs As New ADODB.Recordset
    Dim fila As Integer
    
    On Error GoTo ErrCarga
    Screen.MousePointer = 11
    
    FormateaFlex fg
    Set rs = obj.ListaProveedorSistemaPension()
    If Not rs.EOF Then
        Do While Not rs.EOF
            fg.AdicionaFila
            fila = fg.row
            fg.TextMatrix(fila, 1) = rs!cPersCod
            fg.TextMatrix(fila, 2) = rs!cPersNombre
            fg.TextMatrix(fila, 3) = rs!cTpoSistPens
            fg.TextMatrix(fila, 4) = rs!AFP_cCUSP
            fg.TextMatrix(fila, 5) = rs!AFP_cPersNombre
            fg.TextMatrix(fila, 6) = rs!AFP_cTpoComision
            rs.MoveNext
        Loop
        cargar_datos = True
    End If
    
    RSClose rs
    Set obj = Nothing
    Screen.MousePointer = 0
    Exit Function
ErrCarga:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Function
