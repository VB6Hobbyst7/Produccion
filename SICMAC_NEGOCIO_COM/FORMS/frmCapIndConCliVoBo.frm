VERSION 5.00
Begin VB.Form frmCapIndConCliVoBo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vo Bo Apertura - Concentración de Clientes"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14520
   Icon            =   "frmCapIndConCliVoBo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   14520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SICMACT.FlexEdit FEAperturas 
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   4471
      Cols0           =   10
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Nro-Producto-Titular-TEA-Monto-Indicador Cliente-Id_VoBoConCli-Indicador I-Saldo Cartera-Alerta"
      EncabezadosAnchos=   "500-1200-4000-1200-1200-1500-0-1500-1500-1200"
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-R-R-R-C-R-R-C"
      FormatosEdit    =   "0-0-0-0-4-0-0-0-4-0"
      TextArray0      =   "Nro"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   495
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14295
      Begin VB.ComboBox cboAgencia 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   210
         Width           =   3135
      End
      Begin VB.CommandButton cmdRechazar 
         Caption         =   "&Rechazar"
         Height          =   375
         Left            =   9960
         TabIndex        =   4
         Top             =   160
         Width           =   1215
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   12600
         TabIndex        =   3
         Top             =   160
         Width           =   1215
      End
      Begin VB.CommandButton cmdAprobar 
         Caption         =   "&Aprobar"
         Height          =   375
         Left            =   11280
         TabIndex        =   2
         Top             =   160
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Agencia:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmCapIndConCliVoBo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***************************************************************************************************
'***Nombre      : frmCapIndConCliVoBo
'***Descripción : Formulario para Definir Parametros de los Indicadores de Concetración del Cliente.
'***Creación    : ELRO el 20120920, según OYP-RFC087-2012
'***************************************************************************************************

Private Sub devolverAperturasVoBo()
Dim oNCOMCaptaReportes As New COMNCaptaGenerales.NCOMCaptaReportes
Dim rsAperturas As New ADODB.Recordset
Dim i As Integer

Set rsAperturas = oNCOMCaptaReportes.devolverAperturaVoBo(Trim(Right(cboAgencia.Text, 3)))

LimpiaFlex FEAperturas

If Not rsAperturas.BOF And Not rsAperturas.EOF Then
    FEAperturas.lbEditarFlex = True
    i = 1
    Do While Not rsAperturas.EOF
        FEAperturas.AdicionaFila
        FEAperturas.TextMatrix(i, 1) = rsAperturas!cProducto
        FEAperturas.TextMatrix(i, 2) = PstaNombre(rsAperturas!cPersNombre)
        FEAperturas.TextMatrix(i, 3) = rsAperturas!nTEA
        FEAperturas.TextMatrix(i, 4) = Format(rsAperturas!nMonto, "##,##0.00")
        FEAperturas.TextMatrix(i, 5) = rsAperturas!nIndicadoCliente
        FEAperturas.TextMatrix(i, 6) = rsAperturas!Id_VoBoConCli
        FEAperturas.TextMatrix(i, 7) = rsAperturas!nIndicadoInterno
        FEAperturas.TextMatrix(i, 8) = Format(rsAperturas!nSaldoCartera, "##,###0.00")
        FEAperturas.TextMatrix(i, 9) = rsAperturas!nAlerta
        rsAperturas.MoveNext
        i = i + 1
    Loop
     FEAperturas.lbEditarFlex = False
Else
    MsgBox "No existen Aperturas para ser Rechazados/Aprobados.", vbInformation, "Aviso"
End If

End Sub

Private Sub devolverAgencias()
Dim oDCOMAgencias As New DCOMAgencias
Dim rsAgencias As ADODB.Recordset

Set rsAgencias = oDCOMAgencias.ObtieneAgencias

cboAgencia.Clear

If Not (rsAgencias.BOF And rsAgencias.EOF) Then
    While Not rsAgencias.EOF
        cboAgencia.AddItem rsAgencias!cConsDescripcion & Space(100) & rsAgencias!nConsValor
        rsAgencias.MoveNext
    Wend
End If
cboAgencia.ListIndex = 0
End Sub

Private Sub cboAgencia_Click()
    devolverAperturasVoBo
End Sub

Private Sub cmdAprobar_Click()
Dim lsCliente As String
Dim lnId_VoBoConCli As Long

lsCliente = FEAperturas.TextMatrix(FEAperturas.row, 2)
'lnId_VoBoConCli = CLng(FEAperturas.TextMatrix(FEAperturas.row, 6)) 'RIRO20140710
 lnId_VoBoConCli = CLng(IIf(IsNumeric(FEAperturas.TextMatrix(FEAperturas.row, 6)), FEAperturas.TextMatrix(FEAperturas.row, 6), 0)) 'RIRO20140710

If Len(Trim(FEAperturas.TextMatrix(1, 1))) = 0 And Len(Trim(FEAperturas.TextMatrix(1, 2))) = 0 And FEAperturas.Rows = 2 Then
    MsgBox "No se puede realizar la operación.", vbInformation, "Aviso"
    Exit Sub
End If

If MsgBox("¿Esta seguro que desea Aprobar la Apertura del Cliente " & lsCliente & "?", vbYesNo, "Aviso") = vbYes Then
    Dim oNCOMCaptaReportes As New COMNCaptaGenerales.NCOMCaptaReportes
    Dim oDCOMMov As New COMDMov.DCOMMov
    Dim cMovNroVoBo As String
    Dim lnConfirmar As Long
    
    cMovNroVoBo = oDCOMMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    
    lnConfirmar = oNCOMCaptaReportes.modificarVoBo(lnId_VoBoConCli, 1, cMovNroVoBo)
    If lnConfirmar > 0 Then
        MsgBox "Se Aprobo correctamente.", vbInformation, "Aviso"
         FEAperturas.EliminaFila FEAperturas.row
    Else
        MsgBox "Fallo la Aprobación.", vbInformation, "Aviso"
    End If
    
End If

End Sub

Private Sub cmdRechazar_Click()
Dim lsCliente As String
Dim lnId_VoBoConCli As Long

lsCliente = FEAperturas.TextMatrix(FEAperturas.row, 2)
'lnId_VoBoConCli = CLng(FEAperturas.TextMatrix(FEAperturas.row, 6)) 'RIRO20140710
lnId_VoBoConCli = CLng(IIf(IsNumeric(FEAperturas.TextMatrix(FEAperturas.row, 6)), FEAperturas.TextMatrix(FEAperturas.row, 6), 0)) 'RIRO20140710

If Len(Trim(FEAperturas.TextMatrix(1, 1))) = 0 And Len(Trim(FEAperturas.TextMatrix(1, 2))) = 0 And FEAperturas.Rows = 2 Then
    MsgBox "No se puede realizar la operación.", vbInformation, "Aviso"
    Exit Sub
End If

If MsgBox("¿Esta seguro que desea Rechazar la Apertura del Cliente " & lsCliente & "?", vbInformation + vbYesNo, "Aviso") = vbYes Then
    Dim oNCOMCaptaReportes As New COMNCaptaGenerales.NCOMCaptaReportes
    Dim oDCOMMov As New COMDMov.DCOMMov
    Dim cMovNroVoBo As String
    Dim lnConfirmar As Long

    cMovNroVoBo = oDCOMMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    lnConfirmar = oNCOMCaptaReportes.modificarVoBo(lnId_VoBoConCli, 2, cMovNroVoBo)
    If lnConfirmar > 0 Then
        MsgBox "Se Rechazo correctamente.", vbInformation, "Aviso"
         FEAperturas.EliminaFila FEAperturas.row
    Else
        MsgBox "Fallo el Rechazo.", vbInformation, "Aviso"
    End If
    
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    devolverAgencias
    'devolverAperturasVoBo RIRO20140710
End Sub

