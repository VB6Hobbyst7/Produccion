VERSION 5.00
Begin VB.Form frmParCargosTasa 
   Caption         =   "Parámetro::::Cargos para aprobación de Tasa"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9150
   Icon            =   "frmParCargosTasa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   9150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   320
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txtFecha 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   310
      Left            =   5880
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.ComboBox cboCargos 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   3975
   End
   Begin SICMACT.FlexEdit FECargos 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4048
      Cols0           =   3
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Codigo-Cargo-Movimiento"
      EncabezadosAnchos=   "0-5000-3000"
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
      ColumnasAEditar =   "X-X-X"
      ListaControles  =   "0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-C"
      FormatosEdit    =   "0-0-0"
      TextArray0      =   "Codigo"
      lbUltimaInstancia=   -1  'True
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   3
      Top             =   435
      Width           =   615
   End
   Begin VB.Label lblCargo 
      Caption         =   "Cargo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frmParCargosTasa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Creado por ALPA
'Fecha:::20150205
Option Explicit
Dim lnCantidad As Integer
Private Sub CmdAceptar_Click()
If Len(Trim(Right(cboCargos.Text, 10))) > 0 Then
    Dim sMovNro As String
    Dim objCargos As COMDPersona.DCOMRoles
    Dim ClsMov As COMNContabilidad.NCOMContFunciones
    Set objCargos = New COMDPersona.DCOMRoles
    Set ClsMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = ClsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Call objCargos.GetActualizarDatosParametroPermisoAprobacion(Trim(Right(cboCargos.Text, 10)), sMovNro, 1)
    Call CargarFlexEditCargos
    MsgBox "El dato se guardó satisfactoriamente", vbInformation, "Aviso!"
    Else
        MsgBox "No seleccionó el cargo a guardar", vbInformation, "Aviso!"
    End If
End Sub
Private Sub cmdEliminar_Click()
    If Len(Trim(FECargos.TextMatrix(FECargos.row, 0))) > 0 Then
        Dim sMovNro As String
        Dim objCargos As COMDPersona.DCOMRoles
        Dim ClsMov As COMNContabilidad.NCOMContFunciones
        Set objCargos = New COMDPersona.DCOMRoles
        Set ClsMov = New COMNContabilidad.NCOMContFunciones
        sMovNro = ClsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Call objCargos.GetActualizarDatosParametroPermisoAprobacion(FECargos.TextMatrix(FECargos.row, 0), sMovNro, 0)
        Call CargarFlexEditCargos
        MsgBox "El dato se eliminó satisfactoriamente", vbInformation, "Aviso!"
    Else
        MsgBox "No existe registro que eliminar", vbInformation, "Aviso!"
    End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Call CargarCargos
    Call CargarFlexEditCargos
End Sub
Private Sub CargarCargos()
Dim objConst As COMDConstantes.DCOMConstantes
Set objConst = New COMDConstantes.DCOMConstantes

Dim oRsConst As ADODB.Recordset
Set oRsConst = New ADODB.Recordset
Set oRsConst = objConst.GetCargo
Call Llenar_Combo_con_Recordset(oRsConst, cboCargos)
txtFecha.Text = gdFecSis
End Sub

Private Sub CargarFlexEditCargos()
    Dim objRS As ADODB.Recordset
    Dim objCargos As COMDPersona.DCOMRoles
    Set objCargos = New COMDPersona.DCOMRoles
    FormateaFlex FECargos
    Set objRS = Nothing
    Set objRS = New ADODB.Recordset
    lnCantidad = 0
    Set objRS = objCargos.GetCargarDatosParametroPermisoAprobacion
    Do While Not objRS.EOF
        lnCantidad = lnCantidad + 1
        FECargos.AdicionaFila
        FECargos.TextMatrix(objRS.Bookmark, 0) = objRS!cRHCargoCod
        FECargos.TextMatrix(objRS.Bookmark, 1) = objRS!cRHCargoDescripcion
        FECargos.TextMatrix(objRS.Bookmark, 2) = objRS!cMovNro
        objRS.MoveNext
    Loop
End Sub
