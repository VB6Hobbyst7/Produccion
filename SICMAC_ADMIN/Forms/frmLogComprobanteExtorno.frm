VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogComprobanteExtorno 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extorno"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13095
   Icon            =   "frmLogComprobanteExtorno.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   13095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11880
      TabIndex        =   6
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "&Extornar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   5
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Frame fraGlosa 
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
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   10575
      Begin VB.TextBox txtGlosaExtorno 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   10335
      End
   End
   Begin TabDlg.SSTab SSTabComprobanteExtorno 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Extorno"
      TabPicture(0)   =   "frmLogComprobanteExtorno.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraComprobantesExtorno"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fraComprobantesExtorno 
         Caption         =   "Listado de Comprobantes"
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
         Height          =   3375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   12615
         Begin Sicmact.FlexEdit feComprobanteExtorno 
            Height          =   3015
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   12375
            _ExtentX        =   21828
            _ExtentY        =   5318
            Cols0           =   14
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Proveedor-Tipo-Numero-Emisión-Área Usuaria-Observaciones-Moneda-Monto-nMovNro-cMovNro-cPersCod-nContref-cNContrato"
            EncabezadosAnchos=   "300-2500-1200-1200-1200-1200-2300-1200-1200-0-0-0-0-0"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-C-C-C-C-C-C-C-C-C-C-C-C"
            FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   300
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
   End
End
Attribute VB_Name = "frmLogComprobanteExtorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************
'Nombre : frmLogComprobanteExtorno
'Descripcion:Formulario para el Extorno de Comoprobantes
'Creacion: PASIERS0772014
'*****************************
Option Explicit
Dim gsopecod As String
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
'*******************************

Public Sub Inicio(ByVal psOpeCod As String)
    gsopecod = psOpeCod
    Me.Show 1
End Sub
Private Sub cmdExtornar_Click()
   On Error GoTo ErrCmdExtornar
    If Not validaExtornar Then Exit Sub
         Dim oLog As New NLogGeneral
        Dim oCont As New NContFunciones
        Dim bExito As Boolean
        Dim lnMovNro As Long
        Dim lsMovNro As String
        
        lnMovNro = CLng(feComprobanteExtorno.TextMatrix(feComprobanteExtorno.row, 9))
        If MsgBox("¿Esta seguro de extornar de Comprobante?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
            Exit Sub
        End If
        Screen.MousePointer = 11
            bExito = oLog.ExtornaComprobanteNew(gdFecSis, Right(gsCodAge, 2), gsCodUser, gsopecod, Trim(txtGlosaExtorno.Text), lnMovNro, feComprobanteExtorno.TextMatrix(feComprobanteExtorno.row, 8), feComprobanteExtorno.TextMatrix(feComprobanteExtorno.row, 12), feComprobanteExtorno.TextMatrix(feComprobanteExtorno.row, 13))
        Screen.MousePointer = 0
        If bExito Then
            feComprobanteExtorno.EliminaFila feComprobanteExtorno.row
            txtGlosaExtorno.Text = ""
            MsgBox "Se ha extornado con éxito el Comprobante Nro. " & Trim(feComprobanteExtorno.TextMatrix(feComprobanteExtorno.row, 3)), vbInformation, "Aviso"
            'ARLO 20160126 ***
            Dim lsMonedas As String
            Set objPista = New COMManejador.Pista
            If (gsopecod = 591703 Or gsopecod = 591704) Then
            lsMonedas = "SOLES"
            Else
            lsMonedas = "DOLARES"
            End If
            objPista.InsertarPista gsopecod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", "Se ha extornado con éxito el Comprobante Nro. " & Trim(feComprobanteExtorno.TextMatrix(feComprobanteExtorno.row, 3)) & " en Moneda :" & lsMonedas
            Set objPista = Nothing
            '***
            If MsgBox("¿Desea extornar otro Comprobante?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
            Unload Me
            End If
        Else
            MsgBox "Ha ocurrido un error al extornar el Comprobante Nro. " & Trim(feComprobanteExtorno.TextMatrix(feComprobanteExtorno.row, 3)), vbCritical, "Aviso"
        End If
    Set oLog = Nothing
    Exit Sub
ErrCmdExtornar:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Function validaExtornar() As Boolean
    Dim row As Long
    validaExtornar = True
    If feComprobanteExtorno.TextMatrix(1, 0) = "" Then
        MsgBox "No existen Actas de Conformidad a extornar", vbInformation, "Aviso"
        validaExtornar = False
        Exit Function
    End If
    If Len(Trim(txtGlosaExtorno.Text)) = 0 Then
        MsgBox "Ud. debe de ingresar la Glosa del extorno", vbInformation, "Aviso"
        validaExtornar = False
        txtGlosaExtorno.SetFocus
        Exit Function
    End If
End Function
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub feComprobanteExtorno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtGlosaExtorno.SetFocus
    End If
End Sub
Private Sub Form_Load()
    Select Case gsopecod
        Case gnAlmaComprobanteExtornoMN
            Me.Caption = "Extorno de Registro de Comprobantes en Soles"
        Case gnAlmaComprobanteLibreExtornoMN
            Me.Caption = "Extorno de Registro de Comprobantes Libres en Soles"
        Case gnAlmaComprobanteExtornoME
            Me.Caption = "Extorno de Registro de Comprobantes en Dolares"
        Case gnAlmaComprobanteLibreExtornoME
            Me.Caption = "Extorno de Registro de Comprobantes Libres en Dolares"
    End Select
    CargaDatos
End Sub
Private Sub CargaDatos()
    Dim oLog As New DLogGeneral
    Dim rs As New ADODB.Recordset
    Dim row As Long
    Set rs = oLog.ListaComprobanteExtorno(gsCodUser, gsopecod)
    LimpiaFlex feComprobanteExtorno
    Do While Not rs.EOF
        feComprobanteExtorno.AdicionaFila
        row = feComprobanteExtorno.row
        feComprobanteExtorno.TextMatrix(row, 1) = rs!Proveedor
        feComprobanteExtorno.TextMatrix(row, 2) = rs!Tipo
        feComprobanteExtorno.TextMatrix(row, 3) = rs!Numero
        feComprobanteExtorno.TextMatrix(row, 4) = rs!Emision
        feComprobanteExtorno.TextMatrix(row, 5) = rs!AreaUsuaria
        feComprobanteExtorno.TextMatrix(row, 6) = rs!Observacion
        feComprobanteExtorno.TextMatrix(row, 7) = rs!Moneda
        feComprobanteExtorno.TextMatrix(row, 8) = rs!monto
        feComprobanteExtorno.TextMatrix(row, 9) = rs!nMovNro
        feComprobanteExtorno.TextMatrix(row, 10) = rs!cMovNro
        feComprobanteExtorno.TextMatrix(row, 11) = rs!cPersCod
        feComprobanteExtorno.TextMatrix(row, 12) = rs!nContRef
        feComprobanteExtorno.TextMatrix(row, 13) = rs!cNContrato
        rs.MoveNext
    Loop
    Set rs = Nothing
    Set oLog = Nothing
End Sub
Private Sub txtGlosaExtorno_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        cmdExtornar.SetFocus
    End If
End Sub
