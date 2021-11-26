VERSION 5.00
Begin VB.Form frmMantLiquidezPotencial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Liquidez Potencial"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14055
   Icon            =   "frmMantLiquidezPotencial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   14055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "Quitar"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   12240
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin Sicmact.FlexEdit FEMantLiquidezPotencial 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   7858
      Cols0           =   11
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Nro-Institución Financiera-Moneda-Monto Linea-T.E.A %-Garantía-Fecha Doc-Fecha Ult Act.-Usuario-cMovNro-nNacionalidad"
      EncabezadosAnchos=   "700-4200-800-1800-800-2200-1200-1200-800-0-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "L-L-C-R-R-L-R-R-C-C-R"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-3"
      TextArray0      =   "Nro"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   705
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmMantLiquidezPotencial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************
'*** Nombre : frmMantLiquidezPotencial
'*** Descripción : Formulario para hacer el mentenimiento de Liquidez Potencial.
'*** Creación : MIOL el 20130305, según OYP-ERS025-2013
'********************************************************************************
Option Explicit
Dim lsMovNroAct As String
Dim nMoneda As Integer
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Sub cmdEditar_Click()
 Dim oDInstFinanc As DInstFinanc
 Set oDInstFinanc = New DInstFinanc
 Dim rsLiqPot As ADODB.Recordset
 Set rsLiqPot = New ADODB.Recordset
 
 Set rsLiqPot = oDInstFinanc.ObtieneLiquidezPotencialxMoneda(nMoneda)
 If rsLiqPot.RecordCount > 0 Then
    frmRegLiquidezPotencial.txtInstFinanciera = FEMantLiquidezPotencial.TextMatrix(FEMantLiquidezPotencial.row, 1)
    frmRegLiquidezPotencial.txtInstFinanciera.Enabled = False
    frmRegLiquidezPotencial.txtGarantia = FEMantLiquidezPotencial.TextMatrix(FEMantLiquidezPotencial.row, 5)
    frmRegLiquidezPotencial.txtMontoLinea = Format(FEMantLiquidezPotencial.TextMatrix(FEMantLiquidezPotencial.row, 3), "##,##0.00")
    frmRegLiquidezPotencial.txtFecha = FEMantLiquidezPotencial.TextMatrix(FEMantLiquidezPotencial.row, 6)
    frmRegLiquidezPotencial.txtTEA = FEMantLiquidezPotencial.TextMatrix(FEMantLiquidezPotencial.row, 4)
    frmRegLiquidezPotencial.lblCod = FEMantLiquidezPotencial.TextMatrix(FEMantLiquidezPotencial.row, 9)
    frmRegLiquidezPotencial.chkLinNac = FEMantLiquidezPotencial.TextMatrix(FEMantLiquidezPotencial.row, 10)
    frmRegLiquidezPotencial.Caption = "Edición de Linea"
    frmRegLiquidezPotencial.Show 1
    cargarLineaLiquidezPotencial
    
                    'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Edito la Operación "
                Set objPista = Nothing
                '****
 Else
    MsgBox "No existen datos . . ."
 End If
End Sub

Private Sub cmdQuitar_Click()
 Dim oDInstFinanc As DInstFinanc
 Set oDInstFinanc = New DInstFinanc
  Dim rsLiqPot As ADODB.Recordset
 Set rsLiqPot = New ADODB.Recordset
 
 Set rsLiqPot = oDInstFinanc.ObtieneLiquidezPotencialxMoneda(nMoneda)
 If rsLiqPot.RecordCount > 0 Then
    If MsgBox(" ¿ Seguro de Eliminar Datos ? ", vbOKCancel + vbQuestion, "Confirma grabación") = vbOk Then
       lsMovNroAct = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
    
       Call oDInstFinanc.registrarLiquidezPotencialHist(lsMovNroAct, FEMantLiquidezPotencial.TextMatrix(FEMantLiquidezPotencial.row, 5), FEMantLiquidezPotencial.TextMatrix(FEMantLiquidezPotencial.row, 3), Format(FEMantLiquidezPotencial.TextMatrix(FEMantLiquidezPotencial.row, 6), "yyyyMMdd"), FEMantLiquidezPotencial.TextMatrix(FEMantLiquidezPotencial.row, 4), 3, nMoneda, FEMantLiquidezPotencial.TextMatrix(FEMantLiquidezPotencial.row, 10), FEMantLiquidezPotencial.TextMatrix(FEMantLiquidezPotencial.row, 9))
       Call oDInstFinanc.EliminaLiquidezPotencialxCod(FEMantLiquidezPotencial.TextMatrix(FEMantLiquidezPotencial.row, 9))
       MsgBox "Los datos fueron eliminados Correctamente !!!"
       cargarLineaLiquidezPotencial
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Quito la Operación "
                Set objPista = Nothing
                '****
    End If
  Else
    MsgBox "No existen datos . . ."
 End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
nMoneda = Mid(gsOpeCod, 3, 1)
Me.Caption = "Mantenimiento Liquidez Potencial " + IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME")
CentraForm Me
cargarLineaLiquidezPotencial

End Sub

Private Sub cargarLineaLiquidezPotencial()
 Dim oDInstFinanc As DInstFinanc
 Set oDInstFinanc = New DInstFinanc
 Dim rsLiqPot As ADODB.Recordset
 Set rsLiqPot = New ADODB.Recordset
 Dim i As Integer
    
   Call LimpiaFlex(FEMantLiquidezPotencial)
    
   Set rsLiqPot = oDInstFinanc.ObtieneLiquidezPotencialxMoneda(nMoneda)
        If Not rsLiqPot.BOF And Not rsLiqPot.EOF Then
            i = 1
            FEMantLiquidezPotencial.lbEditarFlex = True
            Do While Not rsLiqPot.EOF
                FEMantLiquidezPotencial.AdicionaFila
                FEMantLiquidezPotencial.TextMatrix(i, 1) = rsLiqPot!cDesc_Institucion
                FEMantLiquidezPotencial.TextMatrix(i, 2) = rsLiqPot!Moneda
                FEMantLiquidezPotencial.TextMatrix(i, 3) = Format(rsLiqPot!nMontoLinea, "##,##0.00")
                FEMantLiquidezPotencial.TextMatrix(i, 4) = Format(rsLiqPot!nTEA, "#0.00")
                FEMantLiquidezPotencial.TextMatrix(i, 5) = rsLiqPot!cGarantia
                FEMantLiquidezPotencial.TextMatrix(i, 6) = Format(rsLiqPot!dFecha, "dd/MM/yyyy")
                FEMantLiquidezPotencial.TextMatrix(i, 7) = rsLiqPot!FechaUltAct
                FEMantLiquidezPotencial.TextMatrix(i, 8) = rsLiqPot!Usuario
                FEMantLiquidezPotencial.TextMatrix(i, 9) = rsLiqPot!cMovNro
                FEMantLiquidezPotencial.TextMatrix(i, 10) = rsLiqPot!nNacionalidad
                i = i + 1
                rsLiqPot.MoveNext
            Loop
        End If
    Set rsLiqPot = Nothing
End Sub
