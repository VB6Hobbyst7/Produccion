VERSION 5.00
Begin VB.Form frmLogActaConformidadExtorno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extorno de Acta de Conformidad"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12810
   Icon            =   "frmLogActaConformidadExtorno.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   12810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Glosa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   60
      TabIndex        =   4
      Top             =   4380
      Width           =   10065
      Begin VB.TextBox txtGlosa 
         Height          =   285
         Left            =   120
         MaxLength       =   300
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   220
         Width           =   9855
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   11475
      TabIndex        =   3
      Top             =   4560
      Width           =   1290
   End
   Begin Sicmact.FlexEdit feActaConformidad 
      Height          =   4335
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   7646
      Cols0           =   10
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-nMovNro-N° de Acta-Tipo-Proveedor-Fecha-Doc. Ref.-Moneda-Monto-sMovNro"
      EncabezadosAnchos=   "400-0-1800-2200-2700-1000-1800-1200-1200-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-L-L-C-L-C-R-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   7
      lbOrdenaCol     =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      Appearance      =   0
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CommandButton cmdExtorno 
      Caption         =   "&Extornar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   10185
      TabIndex        =   2
      Top             =   4560
      Width           =   1290
   End
End
Attribute VB_Name = "frmLogActaConformidadExtorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************
'** Nombre : frmLogActaConformidad
'** Descripción : Registro de Acta de Conformidad creado segun ERS062-2013
'** Creación : EJVG, 20131009 09:00:00 AM
'*************************************************************************
Option Explicit

Private Sub Form_Load()
    CargaControles
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub CargaControles()
    Dim oLog As New DLogGeneral
    Dim rs As New ADODB.Recordset
    Dim row As Long
    Set rs = oLog.ListaActaConformidadXExtorno(gsCodUser)
    LimpiaFlex feActaConformidad
    Do While Not rs.EOF
        feActaConformidad.AdicionaFila
        row = feActaConformidad.row
        feActaConformidad.TextMatrix(row, 1) = rs!nMovNro
        feActaConformidad.TextMatrix(row, 2) = rs!cDocNro
        feActaConformidad.TextMatrix(row, 3) = rs!cDocDesc
        feActaConformidad.TextMatrix(row, 4) = rs!cPersNombre
        feActaConformidad.TextMatrix(row, 5) = Format(rs!dfecha, gsFormatoFechaView)
        feActaConformidad.TextMatrix(row, 6) = rs!cDocReferencia
        feActaConformidad.TextMatrix(row, 7) = rs!cMoneda
        feActaConformidad.TextMatrix(row, 8) = Format(rs!nMovImporte, gsFormatoNumeroView)
        feActaConformidad.TextMatrix(row, 9) = rs!cMovNro
        rs.MoveNext
    Loop
    Set rs = Nothing
    Set oLog = Nothing
End Sub
Private Sub cmdExtorno_Click()
On Error GoTo ErrCmdConforme
    If Not validaExtornar Then Exit Sub
    Dim oLog As New NLogGeneral
    Dim oCont As New NContFunciones
    Dim bExito As Boolean
    Dim lnMovNro As Long
    Dim lsMovNro As String
    
    lnMovNro = CLng(feActaConformidad.TextMatrix(feActaConformidad.row, 1))
    lsMovNro = feActaConformidad.TextMatrix(feActaConformidad.row, 9)
    
    'If Not oCont.PermiteModificarAsiento(lsMovNro) Then Exit Sub 'Comentado PASIERS0772014
    
    If MsgBox("¿Esta seguro de extornar el Acta de Conformidad Digital?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    '''bExito = oLog.ExtornaActaConformidad(gdFecSis, Right(gsCodAge, 2), gsCodUser, IIf(feActaConformidad.TextMatrix(feActaConformidad.row, 7) = "SOLES", gnAlmaActaConformidadExtornoMN, gnAlmaActaConformidadExtornoME), Trim(txtGlosa.Text), lnMovNro) 'MARG ERS044-2016
    bExito = oLog.ExtornaActaConformidad(gdFecSis, Right(gsCodAge, 2), gsCodUser, IIf(feActaConformidad.TextMatrix(feActaConformidad.row, 7) = StrConv(gcPEN_PLURAL, vbUpperCase), gnAlmaActaConformidadExtornoMN, gnAlmaActaConformidadExtornoME), Trim(txtGlosa.Text), lnMovNro) 'MARG ERS044-2016
    Screen.MousePointer = 0
    If bExito Then
        feActaConformidad.EliminaFila feActaConformidad.row
        txtGlosa.Text = ""
        MsgBox "Se ha extornado con éxito el Acta de Conformidad Nro. " & Trim(feActaConformidad.TextMatrix(feActaConformidad.row, 2)), vbInformation, "Aviso"
        If MsgBox("¿Desea extornar otra Acta de Conformidad Digital?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
            Unload Me
        End If
    Else
        MsgBox "Ha ocurrido un error al extornar el Acta de Conformidad Nro. " & Trim(feActaConformidad.TextMatrix(feActaConformidad.row, 3)), vbCritical, "Aviso"
    End If
    Set oLog = Nothing
    Exit Sub
ErrCmdConforme:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub feActaConformidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtGlosa.SetFocus
    End If
End Sub
Private Function validaExtornar() As Boolean
    Dim row As Long
    validaExtornar = True
    If feActaConformidad.TextMatrix(1, 0) = "" Then
        MsgBox "No existen Actas de Conformidad a extornar", vbInformation, "Aviso"
        validaExtornar = False
        Exit Function
    Else
        row = feActaConformidad.row
        
        'Comentado PASI20150422 para permitir extornos en cualquier dia***********************
        '        If DateDiff("D", CDate(feActaConformidad.TextMatrix(row, 5)), gdFecSis) <> 0 Then
        '            MsgBox "La presente Acta de Conformidad no se podra extornar porque solo se permiten extornar las Actas registradas en el día", vbInformation, "Aviso"
        '            validaExtornar = False
        '            Exit Function
        '        End If
        '***********************************************************************
        
    End If
    If Len(Trim(txtGlosa.Text)) = 0 Then
        MsgBox "Ud. debe de ingresar la Glosa del extorno", vbInformation, "Aviso"
        validaExtornar = False
        txtGlosa.SetFocus
        Exit Function
    End If
End Function
Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdExtorno.SetFocus
    End If
End Sub
