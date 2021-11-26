VERSION 5.00
Begin VB.Form frmCapServicioPagoBusqueda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar convenio"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11130
   Icon            =   "frmCapServicioPagoBusqueda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SICMACT.FlexEdit FEConvenios 
      Height          =   1095
      Left            =   2040
      TabIndex        =   5
      Top             =   960
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   1931
      Cols0           =   7
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Empresa-Convenio-cPersCod-IdSerPag-cCodigoConvenio-cCtaCod"
      EncabezadosAnchos=   "500-4000-4000-0-0-0-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   495
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingrese dato a buscar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Width           =   9015
      Begin VB.TextBox txtBusqueda 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   8775
      End
   End
   Begin VB.Frame FRBuscar 
      Caption         =   "Buscar por"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton Option1 
         Caption         =   "Nombre Convenio"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton optForma 
         Caption         =   "Nombre Empresa"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmCapServicioPagoBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'*** Nombre : frmCapServicioPagoBusqueda
'*** Descripción : Formulario para buscar un convenio.
'*** Creación : ELRO el 20130705 11:33:06 AM, según RFC1306270002
'********************************************************************
Option Explicit

Dim fnIdSerPag As Integer
Dim fsNomSerPag As String
Dim fsPersCod As String
Dim fsPersNombre As String
Dim fsCodSerPag As String
Dim fsCtaCod As String 'RIRO20150513 ERS146-2014

Public Sub iniciarBusqueda(ByRef pnIdSerPag As Long, _
                                                ByRef psNomSerPag As String, _
                                                ByRef psPersCod As String, _
                                                ByRef psPersNombre As String, _
                                                ByRef psCodSerPag As String, _
                                                Optional ByRef psCtaCod As String = "")
                                                'RIRO20150513 ERS146-2014, add "psCtaCod"
Show 1
pnIdSerPag = fnIdSerPag
psNomSerPag = fsNomSerPag
psPersCod = fsPersCod
psPersNombre = fsPersNombre
psCodSerPag = fsCodSerPag
psCtaCod = fsCtaCod
End Sub

Private Sub cmdAceptar_Click()
    If Trim(FEConvenios.TextMatrix(1, 0)) = "" Then Exit Sub
    fsPersNombre = FEConvenios.TextMatrix(FEConvenios.row, 1)
    fsNomSerPag = FEConvenios.TextMatrix(FEConvenios.row, 2)
    fsPersCod = FEConvenios.TextMatrix(FEConvenios.row, 3)
    fnIdSerPag = FEConvenios.TextMatrix(FEConvenios.row, 4)
    fsCodSerPag = FEConvenios.TextMatrix(FEConvenios.row, 5)
    fsCtaCod = FEConvenios.TextMatrix(FEConvenios.row, 6) 'RIRO20150513 ERS146-2014
Unload Me
End Sub

Private Sub cmdCancelar_Click()
fsPersNombre = ""
fsNomSerPag = ""
fsPersCod = ""
fnIdSerPag = 0
fsCodSerPag = ""
Unload Me
End Sub

Private Sub FEConvenios_DblClick()
    fsPersNombre = FEConvenios.TextMatrix(FEConvenios.row, 1)
    fsNomSerPag = FEConvenios.TextMatrix(FEConvenios.row, 2)
    fsPersCod = FEConvenios.TextMatrix(FEConvenios.row, 3)
    fnIdSerPag = FEConvenios.TextMatrix(FEConvenios.row, 4)
    fsCodSerPag = FEConvenios.TextMatrix(FEConvenios.row, 5)
    Unload Me
End Sub

Private Sub optForma_Click(Index As Integer)
txtBusqueda.SetFocus
End Sub

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Dim oNCOMCaptaGenerales As NCOMCaptaGenerales
    Set oNCOMCaptaGenerales = New NCOMCaptaGenerales
    Dim rsConvenios As ADODB.Recordset
    Set rsConvenios = New ADODB.Recordset
    LimpiaFlex FEConvenios
    Set rsConvenios = oNCOMCaptaGenerales.buscarConvenioServicioPago(IIf(optForma.iTem(0).value = True, 1, 0), UCase(txtBusqueda))
    If Not (rsConvenios.BOF And rsConvenios.EOF) Then
        'RIRO20150512 ERS146 *****
        If rsConvenios.RecordCount > 0 Then
            cmdAceptar.Default = True
        End If
        'END RIRO ***************
        Do While Not rsConvenios.EOF
            FEConvenios.SetFocus
            FEConvenios.lbEditarFlex = True
            FEConvenios.AdicionaFila
            FEConvenios.TextMatrix(FEConvenios.row, 1) = rsConvenios!cPersNombre
            FEConvenios.TextMatrix(FEConvenios.row, 2) = rsConvenios!cNomSerPag
            FEConvenios.TextMatrix(FEConvenios.row, 3) = rsConvenios!cPersCod
            FEConvenios.TextMatrix(FEConvenios.row, 4) = rsConvenios!Id_SerPag
            FEConvenios.TextMatrix(FEConvenios.row, 5) = rsConvenios!cCodSerPag
            FEConvenios.TextMatrix(FEConvenios.row, 6) = rsConvenios!cCtaCod
            FEConvenios.lbEditarFlex = False
            rsConvenios.MoveNext
        Loop
    Else
        MsgBox "No existe convenio registrado.", vbInformation, "Aviso"
    End If
    Set oNCOMCaptaGenerales = Nothing
    Set rsConvenios = Nothing
Else
    KeyAscii = Letras(KeyAscii)
    cmdAceptar.Default = False
End If
End Sub


