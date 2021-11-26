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
   Begin VB.Frame fraGlosa 
      Caption         =   "Glosa"
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
      Top             =   10
      Width           =   3735
      Begin VB.TextBox txtGlosa 
         Height          =   390
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   225
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
      _ExtentX        =   20849
      _ExtentY        =   7779
      Cols0           =   9
      HighLight       =   2
      AllowUserResizing=   3
      EncabezadosNombres=   "N°-nID-Banco-N° Cheque-Girador-Moneda-Importe-nNroOperacion-nDeposito"
      EncabezadosAnchos=   "350-0-3200-3000-3200-1200-1500-0-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-0-0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-L-L-L-L-C-R-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0"
      TextArray0      =   "N°"
      lbFlexDuplicados=   0   'False
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   3
      lbFormatoCol    =   -1  'True
      lbPuntero       =   -1  'True
      lbOrdenaCol     =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   345
      RowHeight0      =   300
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
            feDetalle.TextMatrix(row, 5) = oRs!cMoneda
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
Private Sub cmdExtornar_Click()
    Dim oCont As COMNContabilidad.NCOMContFunciones
    Dim oDocRec As NCOMDocRec
    Dim fila As Long
    Dim bExito As Boolean
    Dim lsMovNro As String
    Dim lnID As Long
    
    On Error GoTo ErrExtornar
    If feDetalle.TextMatrix(1, 0) = "" Then
        MsgBox "No existen cheques para realizar el extorno", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(Trim(txtGlosa.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la Glosa para continuar", vbInformation, "Aviso"
        If txtGlosa.Visible And txtGlosa.Enabled Then txtGlosa.SetFocus
        Exit Sub
    End If
    fila = feDetalle.row
    If val(feDetalle.TextMatrix(fila, 7)) > 0 Then
        MsgBox "No se puede extornar el Cheque porque ya se realizaron operaciones con el mismo, verifique..", vbExclamation, "Aviso"
        Exit Sub
    End If
    If val(feDetalle.TextMatrix(fila, 8)) > 0 Then
        MsgBox "No se puede extornar el Cheque porque ya realizaron el Depósito en Bancos, verifique..", vbExclamation, "Aviso"
        Exit Sub
    End If
    lnID = CLng(feDetalle.TextMatrix(fila, 1))
    If lnID <= 0 Then
        MsgBox "El Cheque seleccionado no cuenta con un identificador, comuniquese con el Dpto. de TI", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("¿Esta seguro de Extornar el Cheque seleccionado?", vbYesNo + vbInformation, "Aviso") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    Set oDocRec = New NCOMDocRec
    Set oCont = New COMNContabilidad.NCOMContFunciones
    
    lsMovNro = oCont.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    bExito = oDocRec.ExtornarCheque(lnID, lsMovNro, Trim(txtGlosa))
    If bExito Then
        MsgBox "Se ha extornado satisfactoriamente el Cheque", vbInformation, "Aviso"
        feDetalle.EliminaFila fila
        txtGlosa.Text = ""
    Else
        MsgBox "Ha sucedido un error al extornar el Cheque, si el problema persiste comuniquese con el Dpto. de TI", vbInformation, "Aviso"
    End If
    
    Screen.MousePointer = 0
    Set oDocRec = Nothing
    Exit Sub
ErrExtornar:
    Screen.MousePointer = 0
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub txtNroCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdbuscar.Visible And cmdbuscar.Enabled Then cmdbuscar.SetFocus
    End If
End Sub
Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cmdExtornar.Visible And cmdExtornar.Enabled Then cmdExtornar.SetFocus
    End If
End Sub
