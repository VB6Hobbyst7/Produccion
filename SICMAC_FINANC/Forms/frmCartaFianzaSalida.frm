VERSION 5.00
Begin VB.Form frmCartaFianzaSalida 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cartas Fianzas: Salida"
   ClientHeight    =   5385
   ClientLeft      =   945
   ClientTop       =   2040
   ClientWidth     =   9285
   Icon            =   "frmCartaFianzaSalida.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExcel 
      Caption         =   "&Excel"
      Height          =   360
      Left            =   4665
      TabIndex        =   3
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   360
      Left            =   5790
      TabIndex        =   4
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Frame fraCartasFianzas 
      Caption         =   "Lista de Cartas Fianzas"
      Height          =   3210
      Left            =   105
      TabIndex        =   8
      Top             =   600
      Width           =   9045
      Begin Sicmact.FlexEdit fg 
         Height          =   2895
         Left            =   90
         TabIndex        =   1
         Top             =   240
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   5106
         Cols0           =   13
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Nro-N-Itm-Nro Carta-Fec Emision-Persona-Concepto-Entidad Financiera-Monto-Fec Ingreso-Fec Vencim.-Movimiento-nMovNro"
         EncabezadosAnchos=   "400-0-450-1200-1200-3000-3500-3000-1200-1200-1200-1700-0"
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   1
         ListaControles  =   "0-0-4-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-L-C-L-L-L-R-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "Nro"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   6915
      TabIndex        =   5
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   8040
      TabIndex        =   6
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Frame Frame6 
      Caption         =   "&Glosa"
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
      Height          =   945
      Left            =   105
      TabIndex        =   7
      Top             =   3885
      Width           =   9060
      Begin VB.TextBox txtGlosa 
         Height          =   585
         Left            =   180
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   225
         Width           =   8745
      End
   End
   Begin VB.Label lblMoneda 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   330
      Left            =   1335
      TabIndex        =   11
      Top             =   150
      Width           =   705
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Moneda: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   180
      TabIndex        =   10
      Top             =   150
      Width           =   1170
   End
   Begin VB.Label lblFecha 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7800
      TabIndex        =   0
      Top             =   60
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha :"
      Height          =   195
      Left            =   7140
      TabIndex        =   9
      Top             =   105
      Width           =   540
   End
End
Attribute VB_Name = "frmCartaFianzaSalida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsCodCtaDebe As String
Dim lsCodCtaHaber As String
Dim lsTipoDoc As String
Dim lbMN As Boolean
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Sub CargaLista(lsCodCtaCont As String)
Dim rs    As New ADODB.Recordset
Dim nItem  As Integer
Dim oCaja As New nCajaGeneral
On Error GoTo ErrorCargaLista
Set rs = oCaja.GetDatosCartaFianza(lsCodCtaCont)
Set oCaja = Nothing
Do While Not rs.EOF
   fg.AdicionaFila
   nItem = fg.row
   'fg.TextMatrix(nItem, 0) = ""
   fg.TextMatrix(nItem, 1) = nItem
   
   fg.TextMatrix(nItem, 3) = Trim(rs!cDocNro)
   fg.TextMatrix(nItem, 4) = rs!dDocFecha
   fg.TextMatrix(nItem, 5) = PstaNombre(rs!cPersNombre)
   fg.TextMatrix(nItem, 6) = EliminaEnters(Trim(rs!Descripcion))
   fg.TextMatrix(nItem, 7) = Trim(rs!cBanco)
   fg.TextMatrix(nItem, 8) = Format(rs!nMovImporte, gsFormatoNumeroView)
   fg.TextMatrix(nItem, 9) = rs!dFecIng
   
   fg.TextMatrix(nItem, 10) = rs!FECHAVENC
   fg.TextMatrix(nItem, 11) = rs!cMovNro
   fg.TextMatrix(nItem, 12) = rs!nMovNro
   rs.MoveNext
Loop
RSClose rs
Exit Sub
ErrorCargaLista:
   RSClose rs
   MsgBox "Error Nº [" & Err.Number & "] " & TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Function ValidaInterfaz() As Boolean
Dim K As Integer
ValidaInterfaz = False
If Len(Trim(Me.txtGlosa)) = 0 Then
    MsgBox "Ingrese descripción de Operación", vbInformation, "Aviso"
    txtGlosa.SetFocus
    Exit Function
End If
For K = 1 To fg.Rows - 1
   If fg.TextMatrix(K, 2) = "." Then
      ValidaInterfaz = True
      Exit Function
   End If
Next
MsgBox "No se Seleccionó ninguna Carta", vbInformation, "Aviso!"
End Function

Private Sub cmdAceptar_Click()
Dim lsAsiento As String
Dim K As Integer
On Error GoTo ErrorGrabaSalida
If Not ValidaInterfaz Then Exit Sub
If MsgBox(" ¿ Desea Grabar la Salida de Cartas Fianza ? ", vbYesNo + vbQuestion, "Confirmación") Then
   Dim oCaja As New nCajaGeneral
   lsAsiento = oCaja.GrabaCartaFianzaSalida(lblFecha, gsCodAge, gsCodUser, gsOpeCod, txtGlosa, fg.GetRsNew, lsCodCtaDebe, lsCodCtaHaber)
   Set oCaja = Nothing
   If lsAsiento <> "" Then
      EnviaPrevio lsAsiento, "SALIDA DE CARTAS FIANZA", gnLinPage, False
   End If
   fg.EliminaFila fg.row
   If fg.TextMatrix(1, 0) = "" Then
      MsgBox "No existen más Cartas Fianza ", vbInformation, "¡Aviso!"
      Unload Me
   End If
        'ARLO20170217
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Grabo la Operación "
        Set objPista = Nothing
        '****
End If
Exit Sub
ErrorGrabaSalida:
    MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Private Sub cmdExcel_Click()
Dim lsArchivo As String
Dim lbExcel As Boolean
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim oCarta As New nCajaGenImprimir
Dim lsImpre As String
On Error GoTo GeneraEstadError
   lsArchivo = App.path & "\SPOOLER\" & "Cartas_Fianza_" & Format(Date & " " & Time, "yyyymmdd HHMMSS ampm") & "_" & IIf(Mid(gsOpeCod, 3, 1) = gMonedaNacional, "MN", "ME") & ".XLSX"
   lbExcel = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
   If Not lbExcel Then
      Exit Sub
   End If

    Set oCarta = New nCajaGenImprimir
    oCarta.Inicio gsInstCmac, gdFecSis
    lsImpre = oCarta.ExcelCartasFianzaSalida(Me.lblFecha, Mid(gsOpeCod, 3, 1), gnTipCambio, gnLinPage, lsCodCtaDebe, xlAplicacion, xlLibro, xlHoja1, gsNomCmac, gsNomAge)
    If Len(Trim(lsImpre)) = 0 Then
    Else
        ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
        Set oCarta = Nothing
        If lsArchivo <> "" Then
            CargaArchivo lsArchivo, App.path & "\SPOOLER\"
        End If
        MsgBox "Reporte generado satisfactoriamente", vbInformation, "Aviso!!!"
                'ARLO20170217
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Genero Excel "
        Set objPista = Nothing
        '****
    End If
Exit Sub
GeneraEstadError:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
    If lbExcel = True Then
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
    End If
End Sub

Private Sub cmdImprimir_Click()
Dim lnPaginas As Long
Dim lnLineas As Long
Dim I As Integer
Dim Total As Currency
Dim sImpre As String
Dim oImpre As New nCajaGenImprimir
On Error GoTo ImprimirErr
If fg.TextMatrix(1, 0) <> "" Then
   sImpre = oImpre.ImprimeCartasFianzaSalida(Me.lblFecha, Mid(gsOpeCod, 3, 1), gnTipCambio, gnLinPage, lsCodCtaDebe)
   EnviaPrevio sImpre, "REPORTE DE CARTAS FIANZA", gnLinPage, False
        'ARLO20170217
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Realizo Impresión "
        Set objPista = Nothing
        '****
Else
   MsgBox "No Existe Cartas Fianzas en Lista", vbInformation, "Aviso"
End If
Exit Sub
ImprimirErr:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim oOpe As New DOperacion
CentraForm Me
lblFecha = gdFecSis
lsCodCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D")
lsCodCtaHaber = oOpe.EmiteOpeCta(gsOpeCod, "H")
If lsCodCtaDebe <> "" Then
   CargaLista lsCodCtaDebe
End If

lblMoneda.Caption = IIf(Mid(gsOpeCod, 3, 1) = gMonedaNacional, gcMN, gcME)
End Sub

Private Sub Form_Unload(Cancel As Integer)
CierraConexion
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAceptar.SetFocus
End If
End Sub
