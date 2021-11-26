VERSION 5.00
Begin VB.Form frmColPTarifarioCartaNotarialMinka 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  Tarifario de Cartas Notariales - Minka"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   Icon            =   "frmColPTarifarioCartaNotarialMinka.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   5725
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin SICMACT.FlexEdit FETarifario 
         Height          =   4815
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   8493
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Distrito-Tarifa-Codigo"
         EncabezadosAnchos=   "400-4800-800-0"
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
         ColumnasAEditar =   "X-X-2-X"
         ListaControles  =   "0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-C"
         FormatosEdit    =   "0-0-2-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         Enabled         =   0   'False
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   5725
         Width           =   975
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   5520
         TabIndex        =   4
         Top             =   5725
         Width           =   975
      End
      Begin VB.ComboBox cboProvincias 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "Provincia:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmColPTarifarioCartaNotarialMinka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmColPTarifarioCartaNotarialMinka
'** Descripción : Formulario que administra el tarifario de costo por carta notarial en la agencia minka.
'** Creación    : RECO, 20140721 - ERS114-2014
'**********************************************************************************************

Option Explicit

Dim nTipoOperacion As Integer


Public Sub Inicio(ByVal pnOpeTpo As Integer, ByVal psTitulo As String)
    
    nTipoOperacion = 1
    If pnOpeTpo = 1 Then
        Me.Caption = psTitulo & Me.Caption
    End If
    Call CargarProvincias
    cboProvincias.ListIndex = 0
    Me.Show 1
End Sub

Public Sub CargarProvincias()
    Dim oConst As DConstante
    Dim rsProvincias As Recordset
    Dim i As Integer
    
    Set oConst = New DConstante
    Set rsProvincias = New Recordset
    
    Set rsProvincias = oConst.RecuperaConstantes(10037)
    If Not (rsProvincias.EOF And rsProvincias.BOF) Then
        For i = 0 To rsProvincias.RecordCount - 1
            cboProvincias.AddItem "" & rsProvincias!cConsDescripcion
            cboProvincias.ItemData(cboProvincias.NewIndex) = "" & rsProvincias!nconsValor
            rsProvincias.MoveNext
        Next
    End If
End Sub

Private Sub cboProvincias_Click()
    Dim oColp As New COMDColocPig.DCOMColPActualizaBD
    Dim drDirstritos As New ADODB.Recordset
    Dim i As Integer
    
    Set oColp = New COMDColocPig.DCOMColPActualizaBD
    Set drDirstritos = New ADODB.Recordset
    
    Set drDirstritos = oColp.DevuelveDistritoXProvincia(cboProvincias.ItemData(cboProvincias.ListIndex))
    
    If Not (drDirstritos.EOF And drDirstritos.BOF) Then
        FETarifario.Clear
        FormateaFlex FETarifario
        For i = 1 To drDirstritos.RecordCount
            FETarifario.AdicionaFila
            FETarifario.TextMatrix(i, 1) = drDirstritos!cUbiGeoDescripcion
            FETarifario.TextMatrix(i, 2) = Format(drDirstritos!nValor, gcFormView)
            FETarifario.TextMatrix(i, 3) = drDirstritos!cUbiGeoCod
            drDirstritos.MoveNext
        Next
    End If
End Sub

Private Sub cmdCancelar_Click()
    nTipoOperacion = 1
    cmdEditar.Caption = "Editar"
    cmdCancelar.Visible = False
    FETarifario.Enabled = False
End Sub

Private Sub cmdEditar_Click()
    If nTipoOperacion = 1 Then
        nTipoOperacion = 2
        cmdEditar.Caption = "Aceptar"
        cmdCancelar.Visible = True
        FETarifario.Enabled = True
    Else
        Call GuardarCambios
        Call cboProvincias_Click
        nTipoOperacion = 1
        cmdEditar.Caption = "Editar"
        cmdCancelar.Visible = False
        FETarifario.Enabled = False
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Sub GuardarCambios()
    Dim oColp As New COMDColocPig.DCOMColPActualizaBD
    Dim i As Integer
    
    Set oColp = New COMDColocPig.DCOMColPActualizaBD
    
    For i = 1 To FETarifario.Rows - 1
        oColp.ActualizaTarifarioCartanotariaMinka FETarifario.TextMatrix(i, 3), FETarifario.TextMatrix(i, 2)
    Next
End Sub
