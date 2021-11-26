VERSION 5.00
Begin VB.Form frmChequeDepositoExtorno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extorno de Depósito de cheque"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12015
   Icon            =   "frmChequeDepositoExtorno.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   12015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   10800
      TabIndex        =   7
      Top             =   5340
      Width           =   1050
   End
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "&Extornar"
      Height          =   345
      Left            =   9735
      TabIndex        =   6
      Top             =   5340
      Width           =   1050
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   350
      Left            =   5580
      TabIndex        =   5
      Top             =   360
      Width           =   915
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
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
         TabIndex        =   3
         Top             =   220
         Width           =   2235
      End
      Begin VB.Label Label2 
         Caption         =   "Nro. Cheque:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   975
      End
   End
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
      Left            =   8220
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.TextBox txtGlosa 
         Height          =   390
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   225
         Width           =   3510
      End
   End
   Begin SICMACT.FlexEdit feDetalle 
      Height          =   4410
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   7779
      Cols0           =   8
      HighLight       =   2
      AllowUserResizing=   3
      EncabezadosNombres=   "N°-nID-Banco-N° Cheque-Banco Depósito-Cuenta-Moneda-Importe"
      EncabezadosAnchos=   "350-0-2800-2800-2800-2800-1200-1500"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-L-L-L-L-C-R-R"
      FormatosEdit    =   "0-0-0-0-0-0-0-0"
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
      Height          =   480
      Left            =   120
      TabIndex        =   9
      Top             =   5265
      Width           =   11805
   End
End
Attribute VB_Name = "frmChequeDepositoExtorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************
'** Nombre : frmChequeDepositoExtorno
'** Descripción : Para extorno de Cheques creado segun TI-ERS126-2013
'** Creación : EJVG, 20140312 08:36:00 AM
'**********************************************************************
Option Explicit

Private Sub CmdBuscar_Click()
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
    Set oRs = oDocRec.ListaChequexExtornoDeposito(Trim(txtNroCheque.Text), Right(gsCodAge, 2), gdFecSis)
    FormateaFlex feDetalle
    If Not oRs.EOF Then
        Do While Not oRs.EOF
            feDetalle.AdicionaFila
            row = feDetalle.row
            feDetalle.TextMatrix(row, 1) = oRs!nId
            feDetalle.TextMatrix(row, 2) = oRs!cIFiNombre
            feDetalle.TextMatrix(row, 3) = oRs!cNroDoc
            feDetalle.TextMatrix(row, 4) = oRs!cIFiDepositoNombre
            feDetalle.TextMatrix(row, 5) = oRs!cIFiDepositoCtaIFDesc
            feDetalle.TextMatrix(row, 6) = oRs!cMoneda
            feDetalle.TextMatrix(row, 7) = Format(oRs!nMonto, gsFormatoNumeroView)
            oRs.MoveNext
        Loop
    Else
        MsgBox "No se encontraron resultados con los datos ingresados," & Chr(13) & "verifique que el cheque no este valorizado", vbInformation, "Aviso"
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
    Dim oDR As NCOMDocRec
    Dim oCont As NContFunciones
    Dim Fila As Long
    Dim bExito As Boolean
    Dim lsMovNro As String
    Dim lnID As Long
    
    On Error GoTo ErrExtornar
    If feDetalle.TextMatrix(1, 0) = "" Then
        MsgBox "No existen cheques para realizar el extorno", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(Trim(txtglosa.Text)) = 0 Then
        MsgBox "Ud. debe ingresar la Glosa para continuar", vbInformation, "Aviso"
        If txtglosa.Visible And txtglosa.Enabled Then txtglosa.SetFocus
        Exit Sub
    End If
    Fila = feDetalle.row
    lnID = CLng(feDetalle.TextMatrix(Fila, 1))
    If lnID <= 0 Then
        MsgBox "El Cheque seleccionado no cuenta con un identificador, comuniquese con el Dpto. de TI", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("¿Esta seguro de Extornar el Depósito del Cheque seleccionado?", vbYesNo + vbInformation, "Aviso") = vbNo Then Exit Sub
    Set oCont = New NContFunciones
    Set oDR = New NCOMDocRec
    lsMovNro = oCont.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    bExito = oDR.ExtornarDepositoCheque(lnID, lsMovNro, Trim(txtglosa.Text))
    If bExito Then
        MsgBox "Se ha extornado satisfactoriamente el Depósito del cheque", vbInformation, "Aviso"
        feDetalle.EliminaFila Fila
    Else
        MsgBox "Ha sucedido un error al extornar el Depósito del cheque, si el problema persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If
    Set oDR = Nothing
    Set oCont = Nothing
    Exit Sub
ErrExtornar:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub feDetalle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If feDetalle.TextMatrix(1, 0) <> "" Then
            If cmdExtornar.Visible And cmdExtornar.Enabled Then cmdExtornar.SetFocus
        End If
    End If
End Sub
Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cmdExtornar.Visible And cmdExtornar.Enabled Then cmdExtornar.SetFocus
    End If
End Sub
Private Sub txtNroCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CmdBuscar.Visible And CmdBuscar.Enabled Then CmdBuscar.SetFocus
    End If
End Sub
