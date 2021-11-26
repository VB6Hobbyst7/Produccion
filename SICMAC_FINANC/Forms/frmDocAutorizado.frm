VERSION 5.00
Begin VB.Form frmDocAutorizado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Documentos Autorizados SUNAT"
   ClientHeight    =   5025
   ClientLeft      =   330
   ClientTop       =   2175
   ClientWidth     =   16200
   Icon            =   "frmDocAutorizado.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   16200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "&Periodo"
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
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   5340
      Begin VB.CommandButton cmdVer 
         Caption         =   "&Ver"
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
         Left            =   4050
         TabIndex        =   5
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txtAnio 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3060
         MaxLength       =   4
         TabIndex        =   4
         Top             =   270
         Width           =   855
      End
      Begin VB.ComboBox cboMes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmDocAutorizado.frx":030A
         Left            =   570
         List            =   "frmDocAutorizado.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   270
         Width           =   1935
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   2610
         TabIndex        =   3
         Top             =   330
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   315
         Width           =   390
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   14760
      TabIndex        =   11
      Top             =   4440
      Width           =   1305
   End
   Begin VB.Frame Frame2 
      Height          =   585
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   5895
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   345
         Left            =   4280
         TabIndex        =   10
         Top             =   180
         Width           =   1395
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   345
         Left            =   2880
         TabIndex        =   9
         Top             =   180
         Width           =   1395
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Height          =   345
         Left            =   1470
         TabIndex        =   8
         Top             =   180
         Width           =   1395
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   345
         Left            =   60
         TabIndex        =   7
         Top             =   180
         Width           =   1395
      End
   End
   Begin Sicmact.FlexEdit fgReg 
      Height          =   3375
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   15975
      _ExtentX        =   28178
      _ExtentY        =   5953
      Cols0           =   23
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   $"frmDocAutorizado.frx":039A
      EncabezadosAnchos=   "350-1100-450-600-1000-2000-3000-0-3000-1100-1100-1100-0-1100-0-0-0-0-0-0-1200-1200-1200"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-12-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-L-L-L-L-L-L-R-R-R-R-R-C-C-C-L-C-C-R-L-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-2-2-2-0-0-0-0-5-0-2-1-1"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   345
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmDocAutorizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************
'** Nombre : frmDocAutorizado
'** Descripción : Para la visualización de Documentos Autorizados, TI-ERS0532016
'** Creación : PASI, 20161104
'**********************************************************************
Option Explicit
Dim oReg As DRegVenta
Dim dFechaIni As Date
Dim dFechaFin As Date
Dim rs    As New ADODB.Recordset
Private Sub cmdEliminar_Click()
Dim nItem As Integer
Dim ldFecIni As Date, ldFecFin As Date
Dim lcValida As String
Dim lsMovNro As String
Dim oCont As New NContFunciones
    If fgReg.TextMatrix(1, 0) = "" Then Exit Sub
    lcValida = "1"
    If Not oCont.PermiteBorrarRegPorOpe(gsOpeCod, ldFecIni, ldFecFin) Then
        lcValida = "0"
    Else
        If Not (gdFecSis >= Format(ldFecIni, "dd/mm/yyyy") And gdFecSis <= Format(ldFecFin, "dd/mm/yyyy")) Then
            lcValida = "0"
        End If
    End If
    nItem = fgReg.Row
    If lcValida = "0" Then
        If Not oCont.PermiteModificarAsiento(Format(CDate(fgReg.TextMatrix(nItem, 1)), gsFormatoMovFecha), False) Then
            MsgBox "No se puede Eliminar un registro que pertenece a un mes cerrado", vbInformation, "¡Aviso!"
            Set oCont = Nothing
            Exit Sub
        End If
    End If
    Set oCont = Nothing
    If MsgBox(" ¿ Seguro de Eliminar Documento ? ", vbQuestion + vbYesNo + vbDefaultButton1, "!Confirmación!") = vbNo Then Exit Sub
    lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    oReg.EliminaVenta fgReg.TextMatrix(nItem, 2), fgReg.TextMatrix(nItem, 3) & fgReg.TextMatrix(nItem, 4), CDate(fgReg.TextMatrix(nItem, 1)), fgReg.TextMatrix(nItem, 14), lsMovNro
    fgReg.EliminaFila nItem
End Sub
Private Sub cmdImprimir_Click()
If fgReg.TextMatrix(1, 0) = "" Then Exit Sub
ImprimeComprobanteAutorizado fgReg.TextMatrix(fgReg.Row, 3) + fgReg.TextMatrix(fgReg.Row, 4)
End Sub
Private Sub cmdModificar_Click()
If fgReg.TextMatrix(1, 0) = "" Then Exit Sub
dFechaIni = CDate(fgReg.TextMatrix(fgReg.Row, 15))
dFechaFin = gdFecSis
frmDocAutorizadoDet.inicio fgReg.TextMatrix(fgReg.Row, 3) + fgReg.TextMatrix(fgReg.Row, 4)
ListaDatos
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    cboMes.ListIndex = Month(gdFecSis) - 1
    txtAnio = Year(gdFecSis)
    Set oReg = New DRegVenta
End Sub
Private Sub txtAnio_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii): If KeyAscii = 13 And Not Len(txtAnio.Text) = 0 Then cmdVer.SetFocus
End Sub
Private Sub txtAnio_GotFocus()
    fEnfoque txtAnio
End Sub
Private Sub cmdVer_Click()
    ListaDatos
End Sub
Private Sub ListaDatos()
Dim nItem As Long
    If cboMes.ListIndex = -1 Then MsgBox "Falta definir un mes de proceso. Verifique.", vbInformation, "¿Aviso!": cboMes.SetFocus:: Exit Sub
    If Len(txtAnio.Text) = 0 Then MsgBox "Falta definir año de proceso. Verifique.", vbInformation, "¿Aviso!": txtAnio.SetFocus:: Exit Sub
    dFechaIni = CDate("01/" & Format(cboMes.ListIndex + 1, "00") & "/" & Format(txtAnio, "0000"))
    dFechaFin = DateAdd("m", 1, dFechaIni) - 1
    fgReg.Rows = 2: fgReg.EliminaFila 1
    Set rs = oReg.ListaDocAutorizado(dFechaIni, dFechaFin, gsCodArea & gsCodAge)
    If Not (rs.BOF And rs.EOF) Then
        Do While Not rs.EOF
            fgReg.AdicionaFila
            nItem = fgReg.Row
            AsignaValores nItem, rs
            rs.MoveNext
        Loop
    Else
        MsgBox "No existen Comprobantes de Pago en este periodo.", vbOKOnly + vbInformation, "Atención" 'NAGL 20170805
    End If
    RSClose rs
End Sub
Private Sub cmdAgregar_Click()
    frmDocAutorizadoDet.inicio
End Sub
Private Sub AsignaValores(nItem As Long, prs As ADODB.Recordset)
   fgReg.TextMatrix(nItem, 1) = Format(prs!dDocFecha, "dd/mm/yyyy")
   fgReg.TextMatrix(nItem, 2) = Format(prs!nDocTpo, "00")
   fgReg.TextMatrix(nItem, 3) = Format(Mid(prs!cDocNro, 1, 4), "0000") 'NAGL ERS012-2017 Se cambio Long. NroSerie de 3 a 4 Dígitos
   fgReg.TextMatrix(nItem, 4) = Format(Mid(Trim(prs!cDocNro), 5, 20), "0000000") 'NAGL ERS012-2017 Empieza desde la posición 5 y tiene una longitud de 7 Dígitos
   fgReg.TextMatrix(nItem, 5) = IIf(IsNull(prs!cRuc), "", prs!cRuc)
   fgReg.TextMatrix(nItem, 6) = IIf(IsNull(prs!cPersNombre), "", prs!cPersNombre)
   fgReg.TextMatrix(nItem, 7) = prs!cCtaCod
   fgReg.TextMatrix(nItem, 8) = prs!cDescrip
   If prs!nIGV <> 0 Then
      fgReg.TextMatrix(nItem, 9) = Format(prs!nVVenta, gsFormatoNumeroView)
   Else
      fgReg.TextMatrix(nItem, 10) = Format(prs!nVVenta, gsFormatoNumeroView)
   End If
   fgReg.TextMatrix(nItem, 11) = Format(prs!nIGV, gsFormatoNumeroView)
   fgReg.TextMatrix(nItem, 12) = Format(prs!nOtrImp, gsFormatoNumeroView)
   fgReg.TextMatrix(nItem, 13) = Format(prs!nPVenta, gsFormatoNumeroView)
   fgReg.TextMatrix(nItem, 14) = prs!cOpeTpo
   fgReg.TextMatrix(nItem, 15) = Format(rs!dDocFecha, gsFormatoFechaHoraView)
   fgReg.TextMatrix(nItem, 16) = rs!nDocTpo
   fgReg.TextMatrix(nItem, 17) = IIf(IsNull(rs!cDocNroRefe), "", rs!cDocNroRefe)
   fgReg.TextMatrix(nItem, 18) = Format(IIf(IsNull(rs!dDocRefeFec), "", rs!dDocRefeFec), "dd/mm/yyyy")
   fgReg.TextMatrix(nItem, 19) = IIf(prs!cTipoDoc = 2, 6, prs!cTipoDoc)
   If IsNull(prs!nTipoCambio) Then
      fgReg.TextMatrix(nItem, 21) = Format(0, "##,###,##0.000")
   Else
        fgReg.TextMatrix(nItem, 21) = Format(prs!nTipoCambio, "##,###,##0.000")
   End If
   If IsNull(prs!nMoneda) Then
      fgReg.TextMatrix(nItem, 22) = "MN"
   Else
      If prs!nMoneda <> 1 Then
        fgReg.TextMatrix(nItem, 22) = "ME"
        If prs!nTipoCambio <> 0 Then
            fgReg.TextMatrix(nItem, 20) = Format(prs!nPVenta / prs!nTipoCambio, "##,###,##0.000")
        Else
            fgReg.TextMatrix(nItem, 20) = Format(0, "##,###,##0.000")
        End If
      Else
        fgReg.TextMatrix(nItem, 22) = "MN"
      End If
   End If
End Sub
