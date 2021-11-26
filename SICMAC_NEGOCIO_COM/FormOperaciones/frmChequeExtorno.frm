VERSION 5.00
Begin VB.Form frmChequeExtorno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extorno de Cheques Registrados"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11940
   Icon            =   "frmChequeExtorno.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   11940
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmMotExtorno 
      Caption         =   "Motivos del Extorno"
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
      Height          =   2700
      Left            =   4680
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   2845
      Begin VB.CommandButton cmdExtContinuar 
         Caption         =   "&Continuar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   860
         TabIndex        =   13
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtDetExtorno 
         BackColor       =   &H00C0FFFF&
         Height          =   750
         Left            =   240
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox cmbMotivos 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmChequeExtorno.frx":030A
         Left            =   240
         List            =   "frmChequeExtorno.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalles del Extorno"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblExtCmb 
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame fraGlosa 
      Caption         =   "Glosa"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   675
      Left            =   8160
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox txtGlosa 
         BackColor       =   &H80000004&
         Height          =   390
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   225
         Visible         =   0   'False
         Width           =   3510
      End
   End
   Begin VB.Frame fraBuscar 
      Caption         =   "Buscar..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   675
      Left            =   60
      TabIndex        =   7
      Top             =   10
      Width           =   5370
      Begin VB.TextBox txtNroCheque 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   0
         Top             =   220
         Width           =   2235
      End
      Begin VB.Label Label2 
         Caption         =   "Nro. Cheque:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   350
      Left            =   5520
      TabIndex        =   1
      Top             =   250
      Width           =   915
   End
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "&Extornar"
      Height          =   345
      Left            =   9670
      TabIndex        =   4
      Top             =   5235
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   10745
      TabIndex        =   5
      Top             =   5235
      Width           =   1050
   End
   Begin SICMACT.FlexEdit feDetalle 
      Height          =   4410
      Left            =   60
      TabIndex        =   3
      Top             =   735
      Width           =   11820
      _extentx        =   20849
      _extenty        =   7779
      cols0           =   9
      highlight       =   2
      allowuserresizing=   3
      encabezadosnombres=   "N°-nID-Banco-N° Cheque-Girador-Moneda-Importe-nNroOperacion-nDeposito"
      encabezadosanchos=   "350-0-3200-3000-3200-1200-1500-0-0"
      font            =   "frmChequeExtorno.frx":030E
      fontfixed       =   "frmChequeExtorno.frx":0336
      columnasaeditar =   "X-X-X-X-X-X-X-X-X"
      textstylefixed  =   4
      listacontroles  =   "0-0-0-0-0-0-0-0-0"
      encabezadosalineacion=   "C-L-L-L-L-C-R-C-C"
      formatosedit    =   "0-0-0-0-0-0-0-0-0"
      textarray0      =   "N°"
      lbflexduplicados=   0   'False
      lbultimainstancia=   -1  'True
      tipobusqueda    =   3
      lbformatocol    =   -1  'True
      lbpuntero       =   -1  'True
      lbordenacol     =   -1  'True
      lbbuscaduplicadotext=   -1  'True
      colwidth0       =   345
      rowheight0      =   300
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   485
      Left            =   60
      TabIndex        =   6
      Top             =   5160
      Width           =   11805
   End
End
Attribute VB_Name = "frmChequeExtorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************
'** Nombre : frmChequeExtorno
'** Descripción : Para extorno de Cheques creado segun TI-ERS126-2013
'** Creación : EJVG, 20140122 05:51:00 PM
'**********************************************************************
Option Explicit

Private Sub cmdBuscar_Click()
    Dim oDocRec As NCOMDocRec
    Dim oRs As ADODB.Recordset
    Dim row As Long
    
    On Error GoTo ErrBuscar
    If Len(Trim(Me.txtNroCheque.Text)) = 0 Then
        MsgBox "Ud. debe ingresar el Nro. de Cheque para poder realizar la Búsqueda", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    Set oDocRec = New NCOMDocRec
    Set oRs = New ADODB.Recordset
    Set oRs = oDocRec.ListaChequexExtorno(Trim(txtNroCheque.Text), Right(gsCodAge, 2), gdFecSis)
    FormateaFlex feDetalle
    If Not oRs.EOF Then
        Do While Not oRs.EOF
            feDetalle.AdicionaFila
            row = feDetalle.row
            feDetalle.TextMatrix(row, 1) = oRs!nId
            feDetalle.TextMatrix(row, 2) = oRs!cIFiNombre
            feDetalle.TextMatrix(row, 3) = oRs!cNroDoc
            feDetalle.TextMatrix(row, 4) = oRs!cGiradorNombre
            feDetalle.TextMatrix(row, 5) = oRs!cmoneda
            feDetalle.TextMatrix(row, 6) = Format(oRs!nMonto, gsFormatoNumeroView)
            feDetalle.TextMatrix(row, 7) = oRs!nNroOperacion
            feDetalle.TextMatrix(row, 8) = oRs!nDeposito
            oRs.MoveNext
        Loop
    Else
        MsgBox "No se encontraron resultados con los datos ingresados," & Chr(13) & "verifique que el cheque no este depósitado ni valorizado", vbInformation, "Aviso"
    End If
    Screen.MousePointer = 0
    Set oRs = Nothing
    Set oDocRec = Nothing
    Exit Sub
ErrBuscar:
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical, "Aviso"
End Sub

'****CTI3 (ferimoro)  09102018 ****************
'**********************************************
Sub Limpiar()
    frmMotExtorno.Visible = False
    Me.cmbMotivos.ListIndex = -1
    Me.txtDetExtorno.Text = ""
    fraBuscar.Enabled = True
    feDetalle.Enabled = False
    cmdBuscar.Enabled = True
    cmdExtornar.Enabled = False
End Sub
Private Sub cmdExtContinuar_Click()
    Dim oCont As COMNContabilidad.NCOMContFunciones
    Dim oDocRec As NCOMDocRec
    Dim fila As Long
    Dim bExito As Boolean
    Dim lsMovNro As String
    Dim lnId As Long
    
    On Error GoTo ErrExtornar

'    If Len(Trim(txtGlosa.Text)) = 0 Then
'        MsgBox "Ud. debe ingresar la Glosa para continuar", vbInformation, "Aviso"
'        If txtGlosa.Visible And txtGlosa.Enabled Then txtGlosa.SetFocus
'        Exit Sub
'    End If
    If cmbMotivos.ListIndex = -1 Or Len(txtDetExtorno.Text) <= 0 Then
        MsgBox "Debe ingresar el motivo y/o detalle del Extorno", vbInformation, "Aviso"
        Exit Sub
    End If

    '***CTI3 (FERIMORO)   02102018
    Dim DatosExtorna(1) As String
    
    '***CTI3 (ferimoro)    02102018
    frmMotExtorno.Visible = False
    DatosExtorna(0) = cmbMotivos.Text
    DatosExtorna(1) = txtDetExtorno.Text

    fila = feDetalle.row
    If Val(feDetalle.TextMatrix(fila, 7)) > 0 Then
        MsgBox "No se puede extornar el Cheque porque ya se realizaron operaciones con el mismo, verifique..", vbExclamation, "Aviso"
         Call Limpiar
        Exit Sub
    End If
    If Val(feDetalle.TextMatrix(fila, 8)) > 0 Then
        MsgBox "No se puede extornar el Cheque porque ya realizaron el Depósito en Bancos, verifique..", vbExclamation, "Aviso"
        Call Limpiar
        Exit Sub
    End If
    lnId = CLng(feDetalle.TextMatrix(fila, 1))
    If lnId <= 0 Then
        MsgBox "El Cheque seleccionado no cuenta con un identificador, comuniquese con el Dpto. de TI", vbExclamation, "Aviso"
        Call Limpiar
        Exit Sub
    End If
    
    If MsgBox("¿Esta seguro de Extornar el Cheque seleccionado?", vbYesNo + vbInformation, "Aviso") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    Set oDocRec = New NCOMDocRec
    Set oCont = New COMNContabilidad.NCOMContFunciones
    
    lsMovNro = oCont.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    bExito = oDocRec.ExtornarCheque(lnId, lsMovNro, Trim(txtGlosa), DatosExtorna) 'CTI3
    If bExito Then
        MsgBox "Se ha extornado satisfactoriamente el Cheque", vbInformation, "Aviso"
        feDetalle.EliminaFila fila
        txtGlosa.Text = ""
        Call Limpiar
    Else
        MsgBox "Ha sucedido un error al extornar el Cheque, si el problema persiste comuniquese con el Dpto. de TI", vbInformation, "Aviso"
        Call Limpiar
    End If
    
    Screen.MousePointer = 0
    Set oDocRec = Nothing
    Exit Sub
ErrExtornar:
    Screen.MousePointer = 0
End Sub

Private Sub cmdExtornar_Click()
    If feDetalle.TextMatrix(1, 0) = "" Then
        MsgBox "No existen cheques para realizar el extorno", vbInformation, "Aviso"
        Call Limpiar
        Exit Sub
    End If

'******CTI3 (ferimoro) 27092018
 frmMotExtorno.Visible = True
 fraBuscar.Enabled = False
 feDetalle.Enabled = False
 cmdBuscar.Enabled = False
 cmdExtornar.Enabled = False
 cmbMotivos.SetFocus
'******************************
End Sub
'**********************************************
Private Sub cmdsalir_Click()
    Unload Me
End Sub
'******CTI3 (ferimoro) 18102018
Private Sub CargaControles()
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

Set oCons = New COMDConstantes.DCOMConstantes

Set R = oCons.ObtenerConstanteExtornoMotivo

Set oCons = Nothing
Call Llenar_Combo_MotivoExtorno(R, cmbMotivos)

End Sub
Private Sub Form_Load()
Call CargaControles 'CTI3
End Sub

Private Sub txtDetExtorno_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
'If KeyAscii = 13 Then SendKeys "{TAB}": Exit Sub
End Sub

Private Sub txtNroCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdBuscar.Visible And cmdBuscar.Enabled Then cmdBuscar.SetFocus
    End If
End Sub
Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cmdExtornar.Visible And cmdExtornar.Enabled Then cmdExtornar.SetFocus
    End If
End Sub
