VERSION 5.00
Begin VB.Form frmMntCtasContInstancia 
   Caption         =   "Cuenta-Objeto: Definición de Instancias"
   ClientHeight    =   5640
   ClientLeft      =   1380
   ClientTop       =   1635
   ClientWidth     =   8385
   Icon            =   "frmMntCtasContInstancia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8385
   Begin Sicmact.FlexEdit fg 
      Height          =   2535
      Left            =   120
      TabIndex        =   17
      Top             =   2550
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   4471
      Cols0           =   7
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      VisiblePopMenu  =   -1  'True
      EncabezadosNombres=   "#-Cuenta-Ord-Objeto-Descripción-Sub-Uso Age"
      EncabezadosAnchos=   "350-0-450-2200-3400-1400-850"
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
      ColumnasAEditar =   "X-X-X-3-X-5-6"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-1-0-0-4"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-C-L-L-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbFlexDuplicados=   0   'False
      lbUltimaInstancia=   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   345
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   3750
      TabIndex        =   16
      Tag             =   "txtCodigo"
      Top             =   5220
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox txtNombre 
      Height          =   255
      Left            =   3090
      TabIndex        =   15
      Tag             =   "txtNombre"
      Top             =   5220
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5580
      TabIndex        =   2
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   1410
      TabIndex        =   1
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox txtCtaDes 
      BackColor       =   &H00F0FFFF&
      Enabled         =   0   'False
      Height          =   345
      Left            =   1830
      MaxLength       =   255
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   690
      Width           =   6285
   End
   Begin VB.TextBox txtCtaCod 
      BackColor       =   &H00F0FFFF&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   345
      Left            =   1830
      MaxLength       =   20
      TabIndex        =   9
      Top             =   270
      Width           =   1935
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "A&gregar"
      Height          =   375
      Left            =   150
      TabIndex        =   0
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox txtObjetoCod 
      BackColor       =   &H00F0FFFF&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   345
      Left            =   1410
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1335
      Width           =   1935
   End
   Begin VB.TextBox txtObjetoDesc 
      BackColor       =   &H00F0FFFF&
      Enabled         =   0   'False
      Height          =   345
      Left            =   1410
      MaxLength       =   255
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1740
      Width           =   6705
   End
   Begin VB.Label Label7 
      Caption         =   "Instancias de Objeto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   14
      Top             =   2340
      Width           =   2505
   End
   Begin VB.Label Label5 
      Caption         =   "Cuenta Contable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   270
      TabIndex        =   12
      Top             =   330
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   11
      Top             =   750
      Width           =   1155
   End
   Begin VB.Label Label2 
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   7
      Top             =   1800
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Objeto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   5
      Top             =   1410
      Width           =   1155
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1005
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   8145
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Height          =   1065
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   8145
   End
End
Attribute VB_Name = "frmMntCtasContInstancia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql As String
Dim rs   As ADODB.Recordset
Dim sObjCod As String, sDesc As String
Dim sCtaCod As String, sCtaDes As String
Dim lOk As Boolean
Dim nObjNiv As Integer
Dim nCtaObjNiv As Integer
Dim sFiltro As String
Dim sCtaObjOrden As String

Dim clsCtaCont As DCtaCont

Public Sub Inicia(psCta As String, psCtaD As String, psObjCod As String, psDesc As String, pnObjNiv As Integer, pnCtaObjNiv As Integer, psFiltro As String, psCtaObjOrden)
sCtaCod = psCta
sCtaDes = psCtaD
sObjCod = psObjCod
sDesc = psDesc
nObjNiv = pnObjNiv
nCtaObjNiv = pnCtaObjNiv
sFiltro = psFiltro
sCtaObjOrden = psCtaObjOrden
Me.Show 1
End Sub

Private Sub CmdAceptar_Click()
Dim N As Integer
Dim oCaja As DCajaCtasIF

On Error GoTo errAcepta
N = MsgBox(" ¿ Seguro de Aceptar datos ? ", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Confirmación")
If N = vbCancel Then
   Exit Sub
End If
If N = vbNo Then
   Unload Me
End If

Set oCaja = New DCajaCtasIF
Set clsCtaCont = New DCtaCont
clsCtaCont.EliminaCtaObjFiltro txtCtaCod, sCtaObjOrden
Select Case Val(sObjCod)
    Case ObjEntidadesFinancieras
        oCaja.EliminaCtaIFFiltro sCtaCod
    Case ObjCMACAgencias, ObjCMACAgenciaArea, ObjCMACArea
        clsCtaCont.EliminaCtaAreaAgeFiltro sCtaCod, fg.TextMatrix(1, 2)
    Case Else
        clsCtaCont.EliminaCtaObjFiltro txtCtaCod, fg.TextMatrix(1, 2)
End Select

gsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
'clsCtaCont.InsertaCtaObjFiltro txtCtaCod, sCtaObjOrden, txtObjetoCod, "", gsMovNro
For N = 1 To fg.Rows - 1
   If fg.TextMatrix(N, 2) <> "" Then
        Select Case Val(sObjCod)
            Case ObjEntidadesFinancieras
                oCaja.InsertaCtaIFFiltro sCtaCod, Mid(fg.TextMatrix(N, 3), 4, 13), Mid(fg.TextMatrix(N, 3), 1, 2), Mid(fg.TextMatrix(N, 3), 18, 20), fg.TextMatrix(N, 5), IIf(fg.TextMatrix(N, 6) = ".", 1, 0)
            Case ObjCMACAgencias, ObjCMACAgenciaArea, ObjCMACArea
                clsCtaCont.InsertaCtaAreaAgeObjFiltro sCtaCod, fg.TextMatrix(N, 2), txtObjetoCod, Mid(fg.TextMatrix(N, 3), 1, 3), Mid(fg.TextMatrix(N, 3), 4, 2), fg.TextMatrix(N, 5), gsMovNro
            Case Else
              clsCtaCont.InsertaCtaObjFiltro txtCtaCod, fg.TextMatrix(N, 2), fg.TextMatrix(N, 3), fg.TextMatrix(N, 5), gsMovNro
        End Select
   End If
Next
Set oCaja = Nothing
Set clsCtaCont = Nothing

Unload Me
Exit Sub
errAcepta:
    MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Private Sub cmdAgregar_Click()
fg.AdicionaFila
fg.TextMatrix(fg.Row, 1) = txtCtaCod
fg.TextMatrix(fg.Row, 2) = sCtaObjOrden
fg.Col = 3
fg.SetFocus
End Sub

Private Sub CmdCancelar_Click()
If MsgBox(" ¿ Seguro que desea Salir sin Grabar ? ", vbQuestion + vbYesNo, "¡Aviso!") = vbNo Then
    Exit Sub
End If
Unload Me
lOk = False
End Sub

Private Sub cmdEliminar_Click()
fg.EliminaFila fg.Row
fg.SetFocus
End Sub

Private Sub Form_Load()
Dim oFun As New NContFunciones
Dim oAge As New DActualizaDatosArea

CentraForm Me
frmMdiMain.staMain.Panels(2).Text = "Relación : Cuentas Contables - Objetos"
txtCtaCod.Text = sCtaCod
txtCtaDes.Text = sCtaDes
txtObjetoCod.Text = sObjCod
txtObjetoDesc.Text = sDesc
Set clsCtaCont = New DCtaCont
Select Case Val(sObjCod)
    Case ObjEntidadesFinancieras
        Dim oCaja As New DCajaCtasIF
        Set rs = oCaja.CargaCtaIFFiltro(sCtaCod)
        fg.rsTextBuscar = oCaja.CargaCtasIF(Mid(sCtaCod, 3, 1), sFiltro, IIf(nCtaObjNiv > 1, MuestraCuentas, MuestraInstituciones))
        Set oCaja = Nothing
    Case ObjCMACAgencias
        Set rs = clsCtaCont.CargaCtaAreaAgeFiltro(sCtaCod, sObjCod, sFiltro, True)
        fg.rsTextBuscar = oAge.GetAgencias()
    Case ObjCMACAgenciaArea
        Set rs = clsCtaCont.CargaCtaAreaAgeFiltro(sCtaCod, sObjCod, sFiltro, True)
        fg.rsTextBuscar = oAge.GetAgenciasAreas(sFiltro)
    Case ObjCMACArea
        Set rs = clsCtaCont.CargaCtaAreaAgeFiltro(sCtaCod, sObjCod, sFiltro, True)
        fg.rsTextBuscar = oAge.GetAreas()
    Case Else
        Set rs = clsCtaCont.CargaCtaObjFiltro(txtCtaCod, sObjCod, , True)
        fg.rsTextBuscar = oFun.GetObjetosArbol(txtObjetoCod, Trim(sFiltro), nObjNiv + nCtaObjNiv)
End Select
fg.Clear
If Not rs.EOF Then
   Set fg.Recordset = rs
End If
Set oFun = Nothing

fg.lbUltimaInstancia = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmMdiMain.staMain.Panels(2).Text = ""
End Sub

Public Property Get OK() As Integer
OK = lOk
End Property
Public Property Let OK(ByVal vNewValue As Integer)
lOk = OK
End Property

