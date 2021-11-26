VERSION 5.00
Begin VB.Form frmCredListaExoAut 
   Caption         =   "Exoneraciones/Autorizaciones"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6240
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   3795
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin SICMACT.FlexEdit feTiposExonera 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5530
      Cols0           =   5
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Exoneración/Autorización-Tipo-Solicitar-Cod"
      EncabezadosAnchos=   "400-3950-600-1200-0"
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
      ColumnasAEditar =   "X-X-X-3-X"
      ListaControles  =   "0-0-0-4-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-C-C-C"
      FormatosEdit    =   "0-0-0-0-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmCredListaExoAut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmCredListaExoAut
'** Descripción : Formulario para listar las exoneraciones / autorizaciones, con finalidad de aplicar a la sugerencia de credito
'** Creación : RECO, 20140225 - ERS174-2013
'**********************************************************************************************

Option Explicit
Dim oGen As COMDConstSistema.DCOMGeneral
Dim oDNiv As COMDCredito.DCOMNivelAprobacion
Dim oNNiv As COMNCredito.NCOMNivelAprobacion
Dim rs As ADODB.Recordset
Dim lsCtaCod As String
Public pbValSelect As Boolean

Public Sub Inicio(ByVal psCtaCod As String)
    Dim oDNivExo As COMDCredito.DCOMNivelAprobacion
    Set oDNivExo = New COMDCredito.DCOMNivelAprobacion
    ListaTiposExoneracion
    lsCtaCod = psCtaCod
    Call CargarDatos(oDNivExo.ObtieneDatosExoAutCred(lsCtaCod))
    
    oDNivExo.EliminaRegExoneraAutoriza (lsCtaCod)
    Me.Show 1
End Sub

Private Sub ListaTiposExoneracion()
    Dim lnFila As Integer
    Set oDNiv = New COMDCredito.DCOMNivelAprobacion
    Set rs = oDNiv.RecuperaTiposExoneraciones()
    Set oDNiv = Nothing
    Call LimpiaFlex(feTiposExonera)
    If Not rs.EOF Then
        Do While Not rs.EOF
            feTiposExonera.AdicionaFila
            lnFila = feTiposExonera.row
            feTiposExonera.TextMatrix(lnFila, 1) = rs!cExoneraDesc
            feTiposExonera.TextMatrix(lnFila, 2) = IIf(rs!nTipoExoneraCod = 1, "E", "A")
            'feTiposExonera.TextMatrix(lnFila, 3) = rs!cTipoExoneraDesc & Space(25) & rs!nTipoExoneraCod
            feTiposExonera.TextMatrix(lnFila, 4) = rs!cExoneraCod
            rs.MoveNext
        Loop
        feTiposExonera.TopRow = 1
        feTiposExonera.row = 1
    Else
        'cmdNuevoTipos.Enabled = False
        'cmdEliminarTipos.Enabled = False
    End If
End Sub

Private Sub cmdCerrar_Click()
    RegistrarExoneraAutoriza
    Dim i As Integer
    pbValSelect = False
    For i = 1 To Me.feTiposExonera.Rows - 1
        If feTiposExonera.TextMatrix(i, 3) = "." Then
            pbValSelect = True
        End If
    Next
    Unload Me
End Sub

Public Sub RegistrarExoneraAutoriza()
    Dim oDNivExo As COMDCredito.DCOMNivelAprobacion
    Set oDNivExo = New COMDCredito.DCOMNivelAprobacion
    Dim i As Integer
    For i = 1 To Me.feTiposExonera.Rows - 1
        If feTiposExonera.TextMatrix(i, 3) = "." Then
            oDNivExo.RegistraExoneraAutoriza lsCtaCod, Me.feTiposExonera.TextMatrix(i, 4), 1
        End If
    Next
End Sub

Public Sub CargarDatos(poDR As ADODB.Recordset)
    Dim i As Integer
    If Not poDR.EOF Then
        Do While Not poDR.EOF
            For i = 1 To feTiposExonera.Rows - 1
                If (feTiposExonera.TextMatrix(i, 4) = poDR!cExoneraCod) Then
                    feTiposExonera.TextMatrix(i, 3) = "1"
                End If
            Next
            poDR.MoveNext
        Loop
        feTiposExonera.TopRow = 1
        feTiposExonera.row = 1
    Else
        'cmdNuevoTipos.Enabled = False
        'cmdEliminarTipos.Enabled = False
    End If
End Sub
