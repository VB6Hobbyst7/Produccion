VERSION 5.00
Begin VB.Form frmLimitePosCam 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Limites Posición Cambiaria"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3975
   Icon            =   "frmLimitePosCam.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Sicmact.FlexEdit FELimitePosCamb 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2143
      Cols0           =   5
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Nro-Limite-SobreCompra-SobreVenta-nCodLim"
      EncabezadosAnchos=   "0-1200-1200-1200-0"
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
      ColumnasAEditar =   "X-X-2-3-X"
      ListaControles  =   "0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-R-R-C"
      FormatosEdit    =   "0-0-0-0-0"
      TextArray0      =   "Nro"
      lbUltimaInstancia=   -1  'True
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
      CellBackColor   =   -2147483633
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "frmLimitePosCam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'***Nombre:         frmLimitePosCam
'***Descripción:    Formulario que permite registrar Limites.
'***Creación:       MIOL el 20130719 según ERS088-2013 OBJ A
'************************************************************
Option Explicit
Dim oRepCtaColumna As DRepCtaColumna

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdGrabar_Click()
Set oRepCtaColumna = New DRepCtaColumna
Dim nReg As Integer
Dim nInt As Integer
Dim nTem As Integer
Dim nMontoSCReg As Currency
Dim nMontoSCInt As Currency
Dim nMontoSCTem As Currency
Dim nMontoSVReg As Currency
Dim nMontoSVInt As Currency
Dim nMontoSVTem As Currency

    nReg = FELimitePosCamb.TextMatrix(1, 4)
    nInt = FELimitePosCamb.TextMatrix(2, 4)
    nTem = FELimitePosCamb.TextMatrix(3, 4)
    nMontoSCReg = FELimitePosCamb.TextMatrix(1, 2)
    nMontoSCInt = FELimitePosCamb.TextMatrix(2, 2)
    nMontoSCTem = FELimitePosCamb.TextMatrix(3, 2)
    nMontoSVReg = FELimitePosCamb.TextMatrix(1, 3)
    nMontoSVInt = FELimitePosCamb.TextMatrix(2, 3)
    nMontoSVTem = FELimitePosCamb.TextMatrix(3, 3)

    If nMontoSCReg < nMontoSCInt Or nMontoSCReg < nMontoSCTem Or nMontoSCInt < nMontoSCTem Then
        MsgBox "SobreCompra: El limite regulatorio no debe ser menor que el limite interno ni que el limite temprana.", vbInformation, "Verificar"
        Exit Sub
    End If
    If nMontoSVReg < nMontoSVInt Or nMontoSVReg < nMontoSVTem Or nMontoSVInt < nMontoSVTem Then
        MsgBox "SobreVenta: El limite regulatorio no debe ser menor que el limite interno ni que el limite temprana.", vbInformation, "Verificar"
        Exit Sub
    End If

    If MsgBox("  ¿Esta Seguro de Grabar los Montos Ingresados?  ", vbOKCancel, "Confirmación de Asignación") = vbOk Then
        Dim lsMovNro As String
        lsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
        Call oRepCtaColumna.ActualizarLimitePosCamb(nReg, nInt, nTem, nMontoSCReg, nMontoSCInt, nMontoSCTem, nMontoSVReg, nMontoSVInt, nMontoSVTem, lsMovNro)
        Call oRepCtaColumna.InsertarLimitePosCamb(nReg, nInt, nTem, nMontoSCReg, nMontoSCInt, nMontoSCTem, nMontoSVReg, nMontoSVInt, nMontoSVTem, lsMovNro)
        CargarDatos
        MsgBox "Los datos registrados se actualizaron correctamente !!!", vbInformation, "Conforme"
    End If
End Sub

Private Sub FELimitePosCamb_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim lnValor As Currency
    lnValor = CCur(IIf(FELimitePosCamb.TextMatrix(pnRow, pnCol) = "", "0", FELimitePosCamb.TextMatrix(pnRow, pnCol)))
    If lnValor > 100 Or lnValor < 0 Then
        MsgBox " El monto ingresado no debe ser menor que 0 ni mayor que 100! ", vbInformation, "Verificar"
        CargarDatos
        Exit Sub
    End If
    FELimitePosCamb.TextMatrix(pnRow, pnCol) = Format(lnValor, "#,##0.00")
End Sub

Private Sub Form_Load()
 CargarDatos
End Sub

Private Sub CargarDatos()
Set oRepCtaColumna = New DRepCtaColumna
 Dim rsLim As ADODB.Recordset
 Set rsLim = New ADODB.Recordset
 Dim i As Integer

   Call LimpiaFlex(FELimitePosCamb)

   Set rsLim = oRepCtaColumna.GetLimitePosCamb()
        If Not rsLim.BOF And Not rsLim.EOF Then
            i = 1
            FELimitePosCamb.lbEditarFlex = True
            Do While Not rsLim.EOF
                FELimitePosCamb.AdicionaFila
                FELimitePosCamb.TextMatrix(i, 1) = rsLim!Limite
                FELimitePosCamb.TextMatrix(i, 2) = Format(rsLim!SobreCompra, "#,###0.00")
                FELimitePosCamb.TextMatrix(i, 3) = Format(rsLim!SobreVenta, "#,###0.00")
                FELimitePosCamb.TextMatrix(i, 4) = rsLim!Cod
                FELimitePosCamb.BackColor = &H8000000F
                i = i + 1
                rsLim.MoveNext
            Loop
        End If
    Set rsLim = Nothing
    Set oRepCtaColumna = Nothing
End Sub
