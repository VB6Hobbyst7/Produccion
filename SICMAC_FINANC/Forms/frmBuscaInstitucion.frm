VERSION 5.00
Begin VB.Form frmBuscaInstitucion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar Institución"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6495
   Icon            =   "frmBuscaInstitucion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Sicmact.FlexEdit FEInstitucion 
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2990
      Cols0           =   10
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Nro-Institucion-Moneda-MontoLinea-TEA-Garantia-FechaDoc-FechaUlt-Usuario-cMovNro"
      EncabezadosAnchos=   "500-5650-0-0-0-0-0-0-0-0"
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
      EncabezadosAlineacion=   "C-L-L-C-C-L-C-C-L-C"
      FormatosEdit    =   "0-0-0-4-4-0-5-5-0-0"
      TextArray0      =   "Nro"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   495
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   3000
      Width           =   975
   End
   Begin VB.Frame fraInst 
      Caption         =   "Institución"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.TextBox txtInstitucion 
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Top             =   360
         Width           =   5055
      End
   End
End
Attribute VB_Name = "frmBuscaInstitucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nMoneda As Integer

Private Sub cmdAceptar_Click()
 Dim oDInstFinanc As DInstFinanc
 Set oDInstFinanc = New DInstFinanc
 Dim rsLiqPot As ADODB.Recordset
 Set rsLiqPot = New ADODB.Recordset

 Set rsLiqPot = oDInstFinanc.ObtieneLiquidezPotencialxMoneda(nMoneda)
 If rsLiqPot.RecordCount > 0 Then
    frmRegLiquidezPotencial.txtBuscaIFHisManual = FEInstitucion.TextMatrix(FEInstitucion.Row, 1)
    cargarLineaLiquidezPotencial
    Screen.MousePointer = 0
    Unload Me
 Else
    MsgBox "No existen datos . . ."
 End If
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cargarLineaLiquidezPotencial()
 Dim oDInstFinanc As DInstFinanc
 Set oDInstFinanc = New DInstFinanc
 Dim rsLiqPot As ADODB.Recordset
 Set rsLiqPot = New ADODB.Recordset
 Dim i As Integer
    
   Call LimpiaFlex(FEInstitucion)
    
   Set rsLiqPot = oDInstFinanc.ObtieneLiquidezPotencialxMoneda(nMoneda)
        If Not rsLiqPot.BOF And Not rsLiqPot.EOF Then
            i = 1
            FEInstitucion.lbEditarFlex = True
            Do While Not rsLiqPot.EOF
                FEInstitucion.AdicionaFila
                FEInstitucion.TextMatrix(i, 1) = rsLiqPot!cDesc_Institucion
                FEInstitucion.TextMatrix(i, 2) = rsLiqPot!Moneda
                FEInstitucion.TextMatrix(i, 3) = Format(rsLiqPot!nMontoLinea, "##,##0.00")
                FEInstitucion.TextMatrix(i, 4) = Format(rsLiqPot!nTEA, "#0.00")
                FEInstitucion.TextMatrix(i, 5) = rsLiqPot!cGarantia
                FEInstitucion.TextMatrix(i, 6) = Format(rsLiqPot!dFecha, "dd/MM/yyyy")
                FEInstitucion.TextMatrix(i, 7) = rsLiqPot!FechaUltAct
                FEInstitucion.TextMatrix(i, 8) = rsLiqPot!Usuario
                FEInstitucion.TextMatrix(i, 9) = rsLiqPot!cMovNro
                i = i + 1
                rsLiqPot.MoveNext
            Loop
        End If
    Set rsLiqPot = Nothing
End Sub

Private Sub FEInstitucion_Click()
Me.txtInstitucion.Text = FEInstitucion.TextMatrix(FEInstitucion.Row, 1)
End Sub

Private Sub Form_Load()
nMoneda = Mid(gsOpeCod, 3, 1)
cargarLineaLiquidezPotencial
End Sub

Private Sub txtInstitucion_KeyPress(KeyAscii As Integer)
 Dim oDInstFinanc As DInstFinanc
 Set oDInstFinanc = New DInstFinanc
 Dim rsLiqPot As ADODB.Recordset
 Set rsLiqPot = New ADODB.Recordset
 Dim i As Integer

    If KeyAscii = 13 Then
      If Len(Trim(Me.txtInstitucion.Text)) = 0 Then
        MsgBox "Falta Ingresar el Dato", vbInformation, "Aviso"
        Exit Sub
      End If
        Call LimpiaFlex(FEInstitucion)
         
        Set rsLiqPot = oDInstFinanc.ObtieneLiquidezPotencialxDato(Me.txtInstitucion)
             If Not rsLiqPot.BOF And Not rsLiqPot.EOF Then
                 i = 1
                 FEInstitucion.lbEditarFlex = True
                 Do While Not rsLiqPot.EOF
                     FEInstitucion.AdicionaFila
                     FEInstitucion.TextMatrix(i, 1) = rsLiqPot!cDesc_Institucion
                     i = i + 1
                     rsLiqPot.MoveNext
                 Loop
             End If
      If rsLiqPot.RecordCount = 0 Then
        MsgBox "No se Encontraron Datos", vbInformation, "Aviso"
        txtInstitucion.SetFocus
      Else
        FEInstitucion.SetFocus
      End If
      Set rsLiqPot = Nothing
   Else
        KeyAscii = Letras(KeyAscii)
   End If
End Sub
