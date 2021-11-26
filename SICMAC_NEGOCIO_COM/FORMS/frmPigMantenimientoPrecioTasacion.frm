VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPigMantenimientoPrecioTasacion 
   Caption         =   "Configuración de Precio de Tasación"
   ClientHeight    =   2475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   Icon            =   "frmPigMantenimientoPrecioTasacion.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   2475
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin SICMACT.FlexEdit FEValorTasacion 
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2143
      Cols0           =   6
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Tipo cliente-14kt-16kt-18kt-21kt-TasacioCod"
      EncabezadosAnchos=   "2900-900-900-900-900-0"
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
      ColumnasAEditar =   "X-1-2-3-4-X"
      ListaControles  =   "0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "L-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0"
      TextArray0      =   "Tipo cliente"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   2895
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   2990
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Valor de Tasación x Gr. De Oro"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPigMantenimientoPrecioTasacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelar_Click()
    FEValorTasacion.Enabled = False
    FEValorTasacion.lbEditarFlex = False
    cmdCancelar.Visible = False
    cmdGuardar.Visible = False
    CmdEditar.Enabled = True
    
End Sub

Private Sub CmdEditar_Click()
    Dim R As Integer
    Dim c As Integer
    
    R = FEValorTasacion.row
    c = FEValorTasacion.Col
    
    FEValorTasacion.Enabled = True
    FEValorTasacion.lbEditarFlex = True
    cmdCancelar.Visible = True
    cmdGuardar.Visible = True
    CmdEditar.Enabled = False
End Sub

Private Sub Cmdguardar_Click()
    Dim i As Integer
    Dim loPigContrato As New COMDColocPig.DCOMColPContrato
    Set loPigContrato = New COMDColocPig.DCOMColPContrato
    
    Dim nTasacionCod As Integer
    Dim n14kt As Integer
    Dim n16kt As Integer
    Dim n18kt As Integer
    Dim n21kt As Integer
    For i = 1 To FEValorTasacion.rows - 1
        nTasacionCod = FEValorTasacion.TextMatrix(i, 5)
        n14kt = CInt(FEValorTasacion.TextMatrix(i, 1))
        n16kt = FEValorTasacion.TextMatrix(i, 2)
        n18kt = FEValorTasacion.TextMatrix(i, 3)
        n21kt = FEValorTasacion.TextMatrix(i, 4)
    
        loPigContrato.dPigAcutalizaValorTasacion nTasacionCod, n14kt, n16kt, n18kt, n21kt, Format(gdFecSis, "yyyyMMdd") & gsCodUser
    Next
    
    MsgBox "Datos Guardados Correctamente", vbExclamation, "Aviso"
    cmdCancelar_Click
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    inicia
End Sub

Public Sub inicia()
    Dim i As Integer
    Dim nNumFilas As Integer
    Dim loPigContrato  As New COMDColocPig.DCOMColPContrato
    Dim loR As New ADODB.Recordset
    Set loPigContrato = New COMDColocPig.DCOMColPContrato
    Set loR = New ADODB.Recordset
    
    Set loR = loPigContrato.dPigObtenerValoresTasacion
    
    If Not (loR.EOF And loR.BOF) Then
        FEValorTasacion.FormaCabecera
        nNumFilas = loR.RecordCount
        For i = 1 To nNumFilas
            FEValorTasacion.TextMatrix(i, 0) = loR!cConsDescripcion
            FEValorTasacion.TextMatrix(i, 1) = loR!n14kt
            FEValorTasacion.TextMatrix(i, 2) = loR!n16kt
            FEValorTasacion.TextMatrix(i, 3) = loR!n18kt
            FEValorTasacion.TextMatrix(i, 4) = loR!n21kt
            FEValorTasacion.TextMatrix(i, 5) = loR!nTasacionCod
            loR.MoveNext
            If i <> nNumFilas Then
                FEValorTasacion.AdicionaFila
            End If
        Next
    Else
        MsgBox "No se ha encontrado información de la cuenta ingresada"
    End If
    FEValorTasacion.Enabled = False
    FEValorTasacion.lbEditarFlex = False
End Sub
