VERSION 5.00
Begin VB.Form frmCredRiesgoConfCatAge 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Categoria de Agencias"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   Icon            =   "frmCredReisgoConfCatAge.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.TextBox txtAge 
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label lblBuscar 
         Caption         =   "Buscar :"
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
         Left            =   120
         TabIndex        =   5
         Top             =   550
         Width           =   855
      End
   End
   Begin SICMACT.FlexEdit feDatosCatAge 
      Height          =   4860
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   8573
      Cols0           =   11
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-nCodAge-Agencias-nVBajo-Bajo-nVModerado-Moderado-nVAlto-Alto-nVExtremo-Extremo"
      EncabezadosAnchos=   "400-0-2600-0-1000-0-1000-0-1000-0-1000"
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
      EncabezadosAlineacion=   "C-R-L-R-R-R-R-R-R-R-R"
      FormatosEdit    =   "0-3-0-3-2-3-2-3-2-3-2"
      CantEntero      =   10
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmCredRiesgoConfCatAge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsCreditoTotal As ADODB.Recordset

Private Sub CargaDatos()
    
    Dim oCredito As COMDCredito.DCOMCredito
    Dim rsLlenarConfCatAge As ADODB.Recordset
    Dim i As Integer
        
    Set oCredito = New COMDCredito.DCOMCredito
    
    Set rsLlenarConfCatAge = oCredito.LlenarConfCatAge
    
    Set rsCreditoTotal = rsLlenarConfCatAge.Clone
    
    feDatosCatAge.Clear
    feDatosCatAge.FormaCabecera
    Call LimpiaFlex(feDatosCatAge)
    For i = 1 To rsLlenarConfCatAge.RecordCount
        feDatosCatAge.AdicionaFila
        
            feDatosCatAge.TextMatrix(i, 1) = rsLlenarConfCatAge!cAgeCod
            feDatosCatAge.TextMatrix(i, 2) = rsLlenarConfCatAge!cAgeDescripcion
            feDatosCatAge.TextMatrix(i, 3) = rsLlenarConfCatAge!nVBaja
            feDatosCatAge.TextMatrix(i, 4) = Format(rsLlenarConfCatAge!nBaja, "#,#00.00")
            feDatosCatAge.TextMatrix(i, 5) = rsLlenarConfCatAge!nVModerado
            feDatosCatAge.TextMatrix(i, 6) = Format(rsLlenarConfCatAge!nModerado, "#,#00.00")
            feDatosCatAge.TextMatrix(i, 7) = rsLlenarConfCatAge!nVAlto
            feDatosCatAge.TextMatrix(i, 8) = Format(rsLlenarConfCatAge!nAlto, "#,#00.00")
            feDatosCatAge.TextMatrix(i, 9) = rsLlenarConfCatAge!nVExtremo
            feDatosCatAge.TextMatrix(i, 10) = Format(rsLlenarConfCatAge!nExtremo, "#,#00.00")
            
        rsLlenarConfCatAge.MoveNext
    Next i
    
End Sub

Private Sub cmdCancelar_Click()
Call CargaDatos
Dim nVal As Integer
nVal = CDbl(feDatosCatAge.TextMatrix(feDatosCatAge.row, 0))
Select Case CDbl(feDatosCatAge.TextMatrix(feDatosCatAge.row, 0))
    Case nVal
        feDatosCatAge.ColumnasAEditar = "X-X-X-X-X-X-X-X-X-X-X"
    End Select
    
cmdEditar.Enabled = True
cmdGuardar.Enabled = False

End Sub

Private Sub feDatosCatAge_OnCellChange(pnRow As Long, pnCol As Long)
 If IsNumeric(feDatosCatAge.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feDatosCatAge.TextMatrix(pnRow, pnCol) < 0 Then
            feDatosCatAge.TextMatrix(pnRow, pnCol) = 0
        End If
    Else
        feDatosCatAge.TextMatrix(pnRow, pnCol) = 0
    End If
End Sub

Private Sub Form_Load()
    cmdGuardar.Enabled = False
    Call CargaDatos
End Sub

Private Sub CmdEditar_Click()
Dim nValor As Integer
nValor = CDbl(feDatosCatAge.TextMatrix(feDatosCatAge.row, 0))
Select Case CDbl(feDatosCatAge.TextMatrix(feDatosCatAge.row, 0))
    Case nValor
        feDatosCatAge.ColumnasAEditar = "X-X-X-X-5-X-7-X-9-X-11"
        cmdEditar.Enabled = False
        cmdGuardar.Enabled = True
    End Select
End Sub

Private Sub cmdGuardar_Click()
    
    Dim oNCredito As COMNCredito.NCOMCredito
    Dim oDCredito As COMDCredito.DCOMCredito
    Dim rsObtConfMonto As ADODB.Recordset
    Dim MatReferidos As Variant
    Dim i As Integer
        
    Dim GrabarDatos As Boolean

If Not Valida() Then
    Exit Sub
End If

    ReDim MatConfMonto(feDatosCatAge.Rows - 1, 20)
            For i = 1 To feDatosCatAge.Rows - 1
                MatConfMonto(i, 1) = feDatosCatAge.TextMatrix(i, 1)
                
                MatConfMonto(i, 2) = feDatosCatAge.TextMatrix(i, 3)
                MatConfMonto(i, 3) = feDatosCatAge.TextMatrix(i, 4)
                
                MatConfMonto(i, 4) = feDatosCatAge.TextMatrix(i, 5)
                MatConfMonto(i, 5) = feDatosCatAge.TextMatrix(i, 6)
                
                MatConfMonto(i, 6) = feDatosCatAge.TextMatrix(i, 7)
                MatConfMonto(i, 7) = feDatosCatAge.TextMatrix(i, 8)
                
                MatConfMonto(i, 8) = feDatosCatAge.TextMatrix(i, 9)
                MatConfMonto(i, 9) = feDatosCatAge.TextMatrix(i, 10)
            Next i
            
Set oNCredito = New COMNCredito.NCOMCredito

If MsgBox("Los Datos serán Grabados, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub

   GrabarDatos = oNCredito.GrabarDatosConfCatAge(MatConfMonto)
   
    If GrabarDatos = True Then
        MsgBox "Los datos se grabaron Correctamente ?", vbInformation, "Aviso"
        
        feDatosCatAge.ColumnasAEditar = "X-X-X-X-X-X"
            
        cmdEditar.Enabled = True
        cmdGuardar.Enabled = False
        
        Call CargaDatos
    Else
        MsgBox "Hubo error al grabar la informacion", vbError, "Error"
    End If
End Sub

Private Sub txtAge_Change()
    Dim i As Integer
        
    Dim rsFiltro As ADODB.Recordset
    
    Set rsFiltro = rsCreditoTotal.Clone
              
    If Trim(txtAge.Text) <> "" Then
        rsFiltro.Filter = " cAgeDescripcion LIKE '*" + Trim(txtAge.Text) + "*'"
    End If
        
    feDatosCatAge.Clear
    feDatosCatAge.FormaCabecera
    Call LimpiaFlex(feDatosCatAge)
    For i = 1 To rsFiltro.RecordCount
        feDatosCatAge.AdicionaFila

            feDatosCatAge.TextMatrix(i, 1) = rsFiltro!cAgeCod
            feDatosCatAge.TextMatrix(i, 2) = rsFiltro!cAgeDescripcion
            feDatosCatAge.TextMatrix(i, 3) = rsFiltro!nVBaja
            feDatosCatAge.TextMatrix(i, 4) = Format(rsFiltro!nBaja, "#,#00.00")
            feDatosCatAge.TextMatrix(i, 5) = rsFiltro!nVModerado
            feDatosCatAge.TextMatrix(i, 6) = Format(rsFiltro!nModerado, "#,#00.00")
            feDatosCatAge.TextMatrix(i, 7) = rsFiltro!nVAlto
            feDatosCatAge.TextMatrix(i, 8) = Format(rsFiltro!nAlto, "#,#00.00")
            feDatosCatAge.TextMatrix(i, 9) = rsFiltro!nVExtremo
            feDatosCatAge.TextMatrix(i, 10) = Format(rsFiltro!nExtremo, "#,#00.00")

        rsFiltro.MoveNext
    Next i
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 86 And Shift = 2 Then
        KeyCode = 10
    End If
    
    If KeyCode = 113 And Shift = 0 Then
        KeyCode = 10
    End If
End Sub

Private Function Valida() As Boolean
Dim i As Integer
    
    Valida = True

    For i = 1 To (feDatosCatAge.Rows - 1)
        If feDatosCatAge.TextMatrix(i, 4) = 0 Or feDatosCatAge.TextMatrix(i, 6) = 0 Or feDatosCatAge.TextMatrix(i, 8) = 0 Or feDatosCatAge.TextMatrix(i, 10) = 0 Then
            MsgBox "El Monto Asignado debe ser Mayo a Cero, Favor de Verificar", vbInformation, "Aviso"
            Valida = False
            Exit Function
        End If
    Next i
    
End Function
