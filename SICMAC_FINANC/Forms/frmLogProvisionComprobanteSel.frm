VERSION 5.00
Begin VB.Form frmLogProvisionComprobanteSel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección de Doc. Origen"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12780
   Icon            =   "frmLogProvisionComprobanteSel.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   12780
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   10240
      TabIndex        =   1
      Top             =   4680
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   11490
      TabIndex        =   2
      Top             =   4680
      Width           =   1230
   End
   Begin Sicmact.FlexEdit feComprobante 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12555
      _extentx        =   22146
      _extenty        =   7858
      cols0           =   12
      highlight       =   2
      allowuserresizing=   1
      encabezadosnombres=   "#-nMovNro-Tipo Doc.-N° Doc.-F.Emisión-Proveedor-Moneda-Importe-Origen-N° doc. Origen-Glosa-cAbrev"
      encabezadosanchos=   "350-0-1500-1500-1000-2200-900-1200-1200-1600-5000-0"
      font            =   "frmLogProvisionComprobanteSel.frx":030A
      fontfixed       =   "frmLogProvisionComprobanteSel.frx":0336
      columnasaeditar =   "X-X-X-X-X-X-X-X-X-X-X-X"
      textstylefixed  =   3
      listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      encabezadosalineacion=   "C-C-C-L-C-L-C-R-L-C-L-C"
      formatosedit    =   "0-0-0-0-0-0-0-2-0-0-0-0"
      textarray0      =   "#"
      lbultimainstancia=   -1  'True
      lbbuscaduplicadotext=   -1  'True
      colwidth0       =   345
      rowheight0      =   300
   End
End
Attribute VB_Name = "frmLogProvisionComprobanteSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'** Nombre : frmLogProvisionComprobanteSel
'** Descripción : Formulario para la busqueda de Comprobantes segun ERS062-2013
'** Creación : EJVG, 20131113 11:00:00 AM
'******************************************************************************
Dim fRsComprobante As New ADODB.Recordset
Dim fsDocOrigenNro As String
Dim fnComprobMovNro As Long

Option Explicit
Public Sub Inicio(ByVal pRSComprobante As ADODB.Recordset, ByRef psDocOrigenNro As String, ByRef pnComprobMovNro As Long)
    Set fRsComprobante = pRSComprobante.Clone
    Show 1
    psDocOrigenNro = fsDocOrigenNro
    pnComprobMovNro = fnComprobMovNro
End Sub
Private Sub cmdAceptar_Click()
    If feComprobante.TextMatrix(1, 0) = "" Then
        MsgBox "No existen comprobantes pendientes de Provisionar", vbInformation, "Aviso"
        Exit Sub
    End If
    fnComprobMovNro = CLng(feComprobante.TextMatrix(feComprobante.row, 1))
    fsDocOrigenNro = IIf(feComprobante.TextMatrix(feComprobante.row, 11) <> "", feComprobante.TextMatrix(feComprobante.row, 11) & " ", "") & feComprobante.TextMatrix(feComprobante.row, 9)
    Unload Me
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub feComprobante_DblClick()
    If feComprobante.TextMatrix(feComprobante.row, 0) <> "" Then
        cmdAceptar_Click
    End If
End Sub
Private Sub Form_Activate()
    Dim i As Integer
    i = 1
    feComprobante.SetFocus
    SendKeys "{Right}"
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim row As Long
    LimpiaFlex feComprobante
    Do While Not fRsComprobante.EOF
        feComprobante.AdicionaFila
        row = feComprobante.row
        feComprobante.TextMatrix(row, 1) = fRsComprobante!nMovNro
        feComprobante.TextMatrix(row, 2) = fRsComprobante!TpoDoc
        feComprobante.TextMatrix(row, 3) = fRsComprobante!NDoc
        feComprobante.TextMatrix(row, 4) = Format(fRsComprobante!FEmision, gsFormatoFechaView)
        feComprobante.TextMatrix(row, 5) = fRsComprobante!Proveedor
        feComprobante.TextMatrix(row, 6) = fRsComprobante!Moneda
        feComprobante.TextMatrix(row, 7) = Format(fRsComprobante!Importe, gsFormatoNumeroView)
        feComprobante.TextMatrix(row, 8) = fRsComprobante!Origen
        feComprobante.TextMatrix(row, 9) = fRsComprobante!DocOrigen
        feComprobante.TextMatrix(row, 10) = fRsComprobante!Glosa
        feComprobante.TextMatrix(row, 11) = fRsComprobante!cAbrev
        fRsComprobante.MoveNext
    Loop

    If fRsComprobante.RecordCount > 0 Then
        cmdAceptar.Default = True
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set fRsComprobante = Nothing
End Sub
