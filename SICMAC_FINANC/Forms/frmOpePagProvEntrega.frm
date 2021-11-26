VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmOpePagProvEntrega 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Documentos Girados"
   ClientHeight    =   5415
   ClientLeft      =   645
   ClientTop       =   1650
   ClientWidth     =   9765
   Icon            =   "frmOpePagProvEntrega.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   Begin Sicmact.FlexEdit fg 
      Height          =   3645
      Left            =   90
      TabIndex        =   5
      Top             =   720
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   6429
      Cols0           =   13
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Nro.-Ord-Itm-Comprobante-Emisión-Proveedor-cMovDesc-Importe-cPersCod-cMovNro-nMovNro-nDocTpo-cDocNro"
      EncabezadosAnchos=   "450-0-600-2000-1140-4000-0-1200-0-0-0-0-0"
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
      TextStyleFixed  =   3
      ListaControles  =   "0-0-4-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-L-C-L-L-R-L-C-C-C-L"
      FormatosEdit    =   "0-0-0-0-0-0-0-2-0-0-0-0-0"
      TextArray0      =   "Nro."
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbPuntero       =   -1  'True
      lbOrdenaCol     =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   450
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4110
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdReporte 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6150
      TabIndex        =   7
      ToolTipText     =   "Imprime Asientos Contables"
      Top             =   4980
      Width           =   1095
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   450
      Left            =   90
      TabIndex        =   6
      Top             =   4425
      Width           =   9570
   End
   Begin VB.Frame Frame3 
      Caption         =   "Selección"
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
      Height          =   645
      Left            =   7170
      TabIndex        =   11
      Top             =   15
      Width           =   2490
      Begin VB.CommandButton cmdNinguno 
         Caption         =   "&Ninguno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1305
         TabIndex        =   4
         Top             =   255
         Width           =   990
      End
      Begin VB.CommandButton cmdTodos 
         Caption         =   "&Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   3
         Top             =   255
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rango de Fechas"
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
      Height          =   645
      Left            =   90
      TabIndex        =   10
      Top             =   15
      Width           =   5235
      Begin MSMask.MaskEdBox txtFechaDel 
         Height          =   315
         Left            =   795
         TabIndex        =   0
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaAl 
         Height          =   315
         Left            =   2685
         TabIndex        =   1
         Top             =   240
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta :"
         Height          =   195
         Left            =   2085
         TabIndex        =   13
         Top             =   300
         Width           =   510
      End
   End
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
      Height          =   345
      Left            =   8400
      TabIndex        =   9
      Top             =   4980
      Width           =   1095
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "&Confirmar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7275
      TabIndex        =   8
      Top             =   4980
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   240
      Left            =   1515
      TabIndex        =   12
      Top             =   3495
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   423
      _Version        =   393217
      TextRTF         =   $"frmOpePagProvEntrega.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOpePagProvEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsCtaCodB As String
Dim lsCtaCodS As String
Dim lnTpoDoc  As TpoDoc
Dim rs        As ADODB.Recordset
Dim oBarra     As clsProgressBar
Dim WithEvents oImp As nCajaGenImprimir
Attribute oImp.VB_VarHelpID = -1
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Sub cmdConfirmar_Click()
Dim sql As String
Dim i As Integer
Dim lsMov As String
Dim lsMovRef As String
Dim lsReg() As Integer
Dim j As Integer
Dim lbTrans As Boolean
lbTrans = False

On Error GoTo ConfirmarErr

If MsgBox(" ¿ Desea Realizar la Entrega de los Documentos Seleccionados?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmación") = vbNo Then
   Exit Sub
End If

Dim oFun  As New NContFunciones
Dim oCaj As New nCajaGeneral

gsMovNro = oFun.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
If oCaj.GrabaPagoProveedorEntregaDoc(gsMovNro, gsOpeCod, gsOpeDesc, fg.GetRsNew) = 0 Then
   Do While i < fg.Rows
      If fg.TextMatrix(i, 2) = "." Then
         fg.EliminaFila i
      Else
         i = i + 1
      End If
   Loop
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Realizo Operación "
                Set objPista = Nothing
                '****

End If
Set oFun = Nothing
Set oCaj = Nothing

Exit Sub
ConfirmarErr:
    MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Private Sub cmdNinguno_Click()
Marcar 0
End Sub

Private Sub cmdProcesar_Click()
Dim nItem As Integer
Dim lvItem As ListItem
Dim nImporte As Currency, nTipCambio As Currency
Dim rs As ADODB.Recordset
Dim oCaja As New nCajaGeneral

fg.Clear
fg.Rows = 2
fg.FormaCabecera
Set rs = oCaja.GetDatosPagoProveedor(lsCtaCodB & "','" & lsCtaCodS, lnTpoDoc, 0, Format(txtFechaDel, gsFormatoMovFecha), Format(txtFechaAl, gsFormatoMovFecha), False)
If RSVacio(rs) Then
   MsgBox "No existe Comprobantes Pendientes de Entrega", vbInformation, "¡Aviso!"
   Exit Sub
End If
Set oCaja = Nothing

oImp_BarraShow rs.RecordCount
Do While Not rs.EOF
   oImp_BarraProgress rs.Bookmark, "Entrega de Documentos", "", "Procesando...!", vbBlue
   fg.AdicionaFila
   nItem = fg.row
   fg.TextMatrix(nItem, 1) = nItem
   fg.TextMatrix(nItem, 3) = Mid(rs!cDocAbrev & Space(3), 1, 3) & " " & rs!cDocNro
   fg.TextMatrix(nItem, 4) = rs!dDocFecha
   fg.TextMatrix(nItem, 5) = PstaNombre(rs!cPersNombre, False)
   fg.TextMatrix(nItem, 6) = rs!cMovDesc
   fg.TextMatrix(nItem, 7) = Format(rs!nDocImporte * IIf(rs!nDocImporte > 0, 1, -1), gsFormatoNumeroView)
   fg.TextMatrix(nItem, 8) = rs!cPersCod
   fg.TextMatrix(nItem, 9) = rs!cMovNro
   fg.TextMatrix(nItem, 10) = rs!nMovNro
   fg.TextMatrix(nItem, 11) = rs!nDocTpo
   fg.TextMatrix(nItem, 12) = rs!cDocNro
   rs.MoveNext
Loop
RSClose rs
oImp_BarraClose
End Sub

Function Marcar(pValor As Integer)
Dim i As Integer
For i = 1 To fg.Rows - 1
   fg.TextMatrix(i, 2) = pValor
Next
End Function

Private Sub cmdReporte_Click()
Dim lsImpre As String
Set oImp = New nCajaGenImprimir
lsImpre = oImp.ImprimePagoProveedoresEntrega(fg.GetRsNew, gdFecSis, gnLinPage)
EnviaPrevio lsImpre, gsOpeDesc, gnLinPage, False
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Genero Excel "
                Set objPista = Nothing
                '****
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub cmdTodos_Click()
Marcar 1
End Sub

Private Sub Form_Load()
frmOperaciones.Enabled = False
CentraForm Me
txtFechaDel = gdFecSis - 30
txtFechaAl = gdFecSis

Dim oOpe As New DOperacion
lsCtaCodB = oOpe.EmiteOpeCta(gsOpeCod, "D", "0")
lsCtaCodS = oOpe.EmiteOpeCta(gsOpeCod, "D", "1")
lnTpoDoc = oOpe.EmiteDocOpe(gsOpeCod, OpeDocEstObligatorioDebeExistir, OpeDocMetAutogenerado)
Set oOpe = Nothing
RSClose rs

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmOperaciones.Enabled = True
End Sub

Private Sub oImp_BarraClose()
oBarra.CloseForm Me
Set oBarra = Nothing
End Sub

Private Sub oImp_BarraProgress(value As Variant, psTitulo As String, psSubTitulo As String, psTituloBarra As String, ColorLetras As ColorConstants)
oBarra.Progress value, psTitulo, psSubTitulo, psTituloBarra, ColorLetras
End Sub

Private Sub oImp_BarraShow(pnMax As Variant)
Set oBarra = New clsProgressBar
oBarra.ShowForm Me
oBarra.CaptionSyle = eCap_CaptionPercent
oBarra.Max = pnMax
End Sub

Private Sub txtFechaAl_GotFocus()
fEnfoque txtFechaAl
End Sub

Private Sub txtFechaAl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValFecha(txtFechaAl) Then
      cmdProcesar.SetFocus
   End If
End If
End Sub

Private Sub txtFechaDel_GotFocus()
fEnfoque txtFechaDel
End Sub

Private Sub txtFechaDel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValFecha(txtFechaDel) Then
      txtFechaAl.SetFocus
   End If
End If
End Sub
