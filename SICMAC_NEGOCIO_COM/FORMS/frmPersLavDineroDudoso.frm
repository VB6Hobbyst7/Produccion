VERSION 5.00
Begin VB.Form frmPersLavDineroDudoso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Personas Dudosas y Personajes Políticos"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   Icon            =   "frmPersLavDineroDudoso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4665
      TabIndex        =   5
      Top             =   5745
      Width           =   1080
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3525
      TabIndex        =   4
      Top             =   5745
      Width           =   1080
   End
   Begin VB.Frame fraHistoria 
      Caption         =   "Historia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2175
      Left            =   60
      TabIndex        =   3
      Top             =   3480
      Width           =   9210
      Begin VB.CommandButton cmdNuevoEst 
         Caption         =   "&Nuevo Registro"
         Height          =   375
         Left            =   7485
         TabIndex        =   7
         Top             =   1725
         Width           =   1635
      End
      Begin SICMACT.FlexEdit grdHistoria 
         Height          =   1425
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   9000
         _ExtentX        =   15875
         _ExtentY        =   2514
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Fecha-Estado-Documento01-Documento02-Comentario-codigo-smovnro-flag"
         EncabezadosAnchos=   "350-1200-2000-4000-4000-4000-0-0-0"
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
         ColumnasAEditar =   "X-X-2-3-4-5-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-3-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-L-L-L-L-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         Appearance      =   0
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame fraPersona 
      Caption         =   "Personas Registradas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3495
      Left            =   60
      TabIndex        =   0
      Top             =   -15
      Width           =   9210
      Begin VB.CommandButton cmdAgregarPers 
         Caption         =   "&Agregar Persona"
         Height          =   375
         Left            =   7410
         TabIndex        =   2
         Top             =   3000
         Width           =   1695
      End
      Begin SICMACT.FlexEdit grdPersona 
         Height          =   2715
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9000
         _ExtentX        =   15875
         _ExtentY        =   4789
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Codigo-Nombre-Estado-Documento01-Documento02-Comentario-flag"
         EncabezadosAnchos=   "350-1500-3000-1300-4000-4000-4000-0"
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
         ColumnasAEditar =   "X-1-X-3-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-1-0-3-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-L-L-L-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   345
         RowHeight0      =   300
         TipoBusPersona  =   1
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmPersLavDineroDudoso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ClearScreen()
grdPersona.Clear
grdPersona.Rows = 2
grdPersona.FormaCabecera
grdHistoria.Clear
grdHistoria.Rows = 2
grdHistoria.FormaCabecera
ObtienePersonas
End Sub

Private Function ValidaDatosGrid() As Boolean
Dim dFecha As Date
Dim sEstado As String, sFlag As String
Dim i As Integer

'Valida los Nuevos datos de las persona
For i = 1 To grdPersona.Rows - 1
    sFlag = grdPersona.TextMatrix(i, 5)
    If sFlag = "N" Then
        If grdPersona.TextMatrix(i, 1) = "" Or grdPersona.TextMatrix(i, 2) = "" _
            Or grdPersona.TextMatrix(i, 3) = "" Or grdPersona.TextMatrix(i, 4) = "" Then
            MsgBox "Datos ingresados no válidos", vbInformation, "Aviso"
            grdPersona.Row = i
            ValidaDatosGrid = False
            Exit Function
        End If
    End If
Next i
For i = 1 To grdHistoria.Rows - 1
    sFlag = grdHistoria.TextMatrix(i, 5)
    If sFlag = "N" Or sFlag = "M" Then
        If grdHistoria.TextMatrix(i, 1) = "" Or grdHistoria.TextMatrix(i, 2) = "" _
            Or grdHistoria.TextMatrix(i, 3) = "" Or grdHistoria.TextMatrix(i, 4) = "" Then
            MsgBox "Datos ingresados no válidos", vbInformation, "Aviso"
            grdHistoria.Row = i
            ValidaDatosGrid = False
            Exit Function
        End If
    End If
Next i
ValidaDatosGrid = True
End Function

Private Sub cmdAgregarPers_Click()
grdPersona.AdicionaFila
grdPersona.SetFocus
SendKeys "{Enter}"
grdPersona.TextMatrix(grdPersona.Rows - 1, 7) = "N"
'cmdGrabar.Enabled = True
cmdCancelar.Enabled = True
cmdAgregarPers.Enabled = False
cmdNuevoEst.Enabled = True
End Sub

Private Sub cmdCancelar_Click()
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
'cmdNuevoEst.Enabled = True
'cmdComentario.Enabled = True
cmdAgregarPers.Enabled = True
cmdGrabar.Enabled = False

End Sub

Private Sub ObtienePersonas()
Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
Dim rsPers As ADODB.Recordset
Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios

Set rsPers = clsServ.GetPersonasDudLavDinero()
If Not (rsPers.EOF And rsPers.BOF) Then
    Set grdPersona.Recordset = rsPers
    grdPersona_OnRowChange 1, 1
End If

Set clsServ = Nothing
End Sub

Private Sub cmdComentario_Click()
cmdGrabar.Enabled = True
cmdCancelar.Enabled = True
'cmdNuevoEst.Enabled = False
cmdAgregarPers.Enabled = False
grdHistoria.Col = 3
grdHistoria.SetFocus
SendKeys "{Enter}"
grdHistoria.TextMatrix(grdHistoria.Rows - 1, 6) = "M"
End Sub

Private Sub cmdGrabar_Click()
If Not ValidaDatosGrid Then Exit Sub

If MsgBox("¿Desea grabar la información?", vbQuestion + vbYesNo, "Grabar") = vbYes Then
        
    Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
    Dim sMovNro As String
    Dim clsMov As COMNContabilidad.NCOMContFunciones
    Dim rsPers As ADODB.Recordset, rsHist As ADODB.Recordset
    
    Set clsMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set clsMov = Nothing
    Set rsPers = grdPersona.GetRsNew()
    Set rsHist = grdHistoria.GetRsNew()
    
    Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
    clsServ.ActualizaPersDudLavDinero rsPers, rsHist, sMovNro
    Set clsServ = Nothing
    ClearScreen
End If
End Sub

Private Sub cmdNuevoEst_Click()
cmdGrabar.Enabled = True
cmdCancelar.Enabled = True
'cmdNuevoEst.Enabled = False
'cmdComentario.Enabled = False
cmdAgregarPers.Enabled = False
grdHistoria.AdicionaFila
grdHistoria.SetFocus
SendKeys "{Enter}"
grdHistoria.TextMatrix(grdHistoria.Rows - 1, 1) = Format$(gdFecSis, gcFormatoFechaView)
grdHistoria.TextMatrix(grdHistoria.Rows - 1, 8) = "N"
grdHistoria.TextMatrix(grdHistoria.Rows - 1, 6) = grdPersona.TextMatrix(grdPersona.Row, 1)
End Sub

'By Capi 28012008 modificacion general por ser formulario nuevo
Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
Me.Caption = "Registro Personas Dudosas del Lavado Dinero"
Dim rsEst As ADODB.Recordset
Dim rsHist As ADODB.Recordset
Set rsHist = New ADODB.Recordset

Dim clsGen As COMDConstSistema.DCOMGeneral
Set clsGen = New COMDConstSistema.DCOMGeneral
Set rsEst = clsGen.GetConstante(gPersCondControl)
Set rsHist = rsEst.Clone
grdPersona.CargaCombo rsEst
'Set rsEst = clsGen.GetConstante(gPersEstLavDinero)
grdHistoria.CargaCombo rsHist 'rsEst
Set clsGen = Nothing
ObtienePersonas
cmdNuevoEst.Enabled = True
'cmdComentario.Enabled = False
End Sub

Private Sub grdPersona_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
If psDataCod = "" Then
    grdPersona.EliminaFila pnRow
End If
If pbEsDuplicado Then
    MsgBox "Persona registrada. Ingrese un nuevo estado.", vbInformation, "Aviso"
    grdPersona.EliminaFila pnRow
    If cmdNuevoEst.Enabled Then cmdNuevoEst.SetFocus
End If
End Sub

Private Sub grdPersona_OnRowChange(pnRow As Long, pnCol As Long)
Dim rsHist As ADODB.Recordset
Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
Dim sPersCod As String

sPersCod = grdPersona.TextMatrix(pnRow, 1)

Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
Set rsHist = clsServ.GetPersonaHistDudLavDinero(sPersCod)
Set clsServ = Nothing

If Not (rsHist.EOF And rsHist.BOF) Then
    Set grdHistoria.Recordset = rsHist
    'cmdNuevoEst.Enabled = True
    'cmdComentario.Enabled = True
Else
    grdHistoria.Clear
    grdHistoria.Rows = 2
    grdHistoria.FormaCabecera
    'cmdNuevoEst.Enabled = False
    'cmdComentario.Enabled = False
End If
Set rsHist = Nothing
End Sub

