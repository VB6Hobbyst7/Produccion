VERSION 5.00
Begin VB.Form frmParametrosPIT 
   Caption         =   "Parametros PIT"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4620
      TabIndex        =   4
      Top             =   4560
      Width           =   990
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6780
      TabIndex        =   3
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   5700
      TabIndex        =   2
      Top             =   4560
      Width           =   975
   End
   Begin VB.Frame fraParam 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7650
      Begin SICMACT.FlexEdit grdParam 
         Height          =   4005
         Left            =   105
         TabIndex        =   1
         Top             =   210
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   7064
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Constante-Descripción-Moneda-Valor-Tag"
         EncabezadosAnchos=   "350-0-4500-1200-1000-0"
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
         ColumnasAEditar =   "X-X-X-X-3-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-L-L-R"
         FormatosEdit    =   "0-0-0-0-2-0"
         AvanceCeldas    =   1
         TextArray0      =   "#"
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
      End
   End
End
Attribute VB_Name = "frmParametrosPIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'By capi 21012009
Dim objPista As COMManejador.Pista
Dim nmoneda As COMDConstantes.Moneda

Public Sub inicia(Optional bConsulta As Boolean = False)
If bConsulta Then
    cmdCancelar.Visible = False
    cmdGrabar.Visible = False
    cmdEditar.Visible = False
    Me.Caption = "Captaciones - Parámetros - Consulta"
    grdParam.lbEditarFlex = False
Else
    cmdCancelar.Visible = True
    cmdGrabar.Visible = True
    cmdEditar.Visible = True
    cmdCancelar.Enabled = False
    cmdGrabar.Enabled = False
    Me.Caption = "Captaciones - Parámetros - Mantenimiento"
End If
Me.Show 1
End Sub

Private Sub cmdCancelar_Click()
cmdEditar.Enabled = True
cmdCancelar.Enabled = False
cmdGrabar.Enabled = False
cmdSalir.Enabled = True
grdParam.SetFocus
End Sub

Private Sub cmdEditar_Click()
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

For i = 1 To grdParam.Rows - 1
    If grdParam.TextMatrix(i, 5) = "M" Then
        clsparam.ActualizaParametrosPIT grdParam.TextMatrix(i, 1), grdParam.TextMatrix(i, 2), CDbl(grdParam.TextMatrix(i, 4))
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


Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim clsparam  As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim rsPar As ADODB.Recordset
    Set clsparam = New COMNCaptaGenerales.NCOMCaptaDefinicion
    Set rsPar = New ADODB.Recordset
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    rsPar.CursorLocation = adUseClient
    Set rsPar = clsparam.GetParametrosPit
    grdParam.rsFlex = rsPar
    Set rsPar = Nothing
    Set clsparam = Nothing
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCapMantParametros
End Sub

Private Sub grdParam_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmdGrabar.Visible Then
        cmdEditar_Click
    End If
    
    If grdParam.Col = 3 Then
        If grdParam.TextMatrix(grdParam.Row, 4) < 0 Then
            MsgBox "Ingrese solo valores positivos", vbInformation, "Aviso"
        End If
    End If

End If
End Sub

Private Sub grdParam_OnRowChange(pnRow As Long, pnCol As Long)
grdParam.TextMatrix(pnRow, 5) = "M"
If grdParam.Col = 3 Then
    If grdParam.TextMatrix(grdParam.Row, 4) < 0 Then
        MsgBox "Ingrese solo valores positivos", vbInformation, "Aviso"
    End If
End If
End Sub

Private Sub grdParam_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    If grdParam.Col = 4 Then
    If grdParam.TextMatrix(grdParam.Row, 4) < 0 Then
        MsgBox "Ingrese solo valores positivos", vbInformation, "Aviso"
    End If
End If

End Sub
