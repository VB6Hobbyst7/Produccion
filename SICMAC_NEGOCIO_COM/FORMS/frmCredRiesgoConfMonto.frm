VERSION 5.00
Begin VB.Form frmCredRiesgoConfMonto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuracion de Monto"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   Icon            =   "frmCredRiesgoConfMonto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   4650
      TabIndex        =   3
      Top             =   0
      Width           =   2415
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Mora Cosecha No Aceptable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   2240
      TabIndex        =   2
      Top             =   0
      Width           =   2430
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Mora Cosecha Aceptable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   6045
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
   Begin SICMACT.FlexEdit feConfMonto 
      Height          =   1620
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   2858
      Cols0           =   7
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Nivel Riesgo Agencia-Riesgo 1-Riesgo 2-Riesgo 1-Riesgo 2-nCodNivRiesgo"
      EncabezadosAnchos=   "400-1700-1200-1200-1200-1200-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-R-R-R-R-R"
      FormatosEdit    =   "0-0-2-2-2-2-3"
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
Attribute VB_Name = "frmCredRiesgoConfMonto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim Aceptable As Integer
    Dim NoAceptable As Integer
    Dim RA1 As Integer
    Dim RA2 As Integer
    Dim RNA1 As Integer
    Dim RNA2 As Integer
    Dim fnTipoPermiso As Integer

Private Sub cmdCancelar_Click()
Call LLenarGrilla

Select Case CInt(feConfMonto.TextMatrix(feConfMonto.row, 0))
    Case 1, 2, 3, 4
        feConfMonto.ColumnasAEditar = "X-X-X-X-X"
    End Select
    
cmdEditar.Enabled = True
cmdGuardar.Enabled = False

End Sub

Private Sub CmdEditar_Click()
Select Case CDbl(feConfMonto.TextMatrix(feConfMonto.row, 0))
    Case 1, 2, 3, 4
        feConfMonto.ColumnasAEditar = "X-X-3-4-5"
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
    ReDim MatConfMonto(feConfMonto.Rows - 1, 12)
            For i = 1 To feConfMonto.Rows - 1
                MatConfMonto(i, 1) = feConfMonto.TextMatrix(i, 1)
                MatConfMonto(i, 2) = Aceptable
                MatConfMonto(i, 3) = RA1
                MatConfMonto(i, 4) = feConfMonto.TextMatrix(i, 2)
                MatConfMonto(i, 5) = RA2
                MatConfMonto(i, 6) = feConfMonto.TextMatrix(i, 3)
                MatConfMonto(i, 7) = NoAceptable
                MatConfMonto(i, 8) = RNA1
                MatConfMonto(i, 9) = feConfMonto.TextMatrix(i, 4)
                MatConfMonto(i, 10) = RNA2
                MatConfMonto(i, 11) = feConfMonto.TextMatrix(i, 5)
                MatConfMonto(i, 12) = feConfMonto.TextMatrix(i, 6)
            Next i
            
Set oNCredito = New COMNCredito.NCOMCredito

If MsgBox("Los Datos serán Grabados, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub

   GrabarDatos = oNCredito.GrabarDatosConfMonto(MatConfMonto)
   
    If GrabarDatos = True Then
        MsgBox "Los datos se grabaron Correctamente ?", vbInformation, "Aviso"
        
        Select Case CInt(feConfMonto.TextMatrix(feConfMonto.row, 0))
            Case 1, 2, 3, 4
                feConfMonto.ColumnasAEditar = "X-X-X-X-X"
            End Select
            
        cmdEditar.Enabled = True
        cmdGuardar.Enabled = False
        
        Call LLenarGrilla
    Else
        MsgBox "Hubo error al grabar la informacion", vbError, "Error"
    End If
End Sub

Private Function Valida() As Boolean
Dim i As Integer
    
    Valida = True

    For i = 1 To (feConfMonto.Rows - 1)
        If feConfMonto.TextMatrix(i, 2) = 0 Or feConfMonto.TextMatrix(i, 3) = 0 Or feConfMonto.TextMatrix(i, 4) = 0 Or feConfMonto.TextMatrix(i, 5) = 0 Then
            MsgBox "El Monto Asignado al Riesgo debe ser Mayo a Cero, Favor de Verificar", vbInformation, "Aviso"
            Valida = False
            Exit Function
        End If
    Next i
    
End Function

Private Sub feConfMonto_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(feConfMonto.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feConfMonto.TextMatrix(pnRow, pnCol) < 0 Then
            feConfMonto.TextMatrix(pnRow, pnCol) = 0
        End If
    Else
        feConfMonto.TextMatrix(pnRow, pnCol) = 0
    End If
End Sub

Private Sub Form_Load()
    Call LLenarGrilla
    cmdGuardar.Enabled = False
End Sub

Public Sub LLenarGrilla()

    Dim oCredito As COMDCredito.DCOMCredito
    Dim rsLlenarConfMont As ADODB.Recordset
    Dim i As Integer
        
    Set oCredito = New COMDCredito.DCOMCredito
    
    Set rsLlenarConfMont = oCredito.LlenarConfMonto
    
    feConfMonto.Clear
    feConfMonto.FormaCabecera
    Call LimpiaFlex(feConfMonto)
    For i = 1 To rsLlenarConfMont.RecordCount
        feConfMonto.AdicionaFila
            feConfMonto.TextMatrix(i, 1) = rsLlenarConfMont!NR
            Aceptable = rsLlenarConfMont!nA
            RA1 = rsLlenarConfMont!nRA1
            feConfMonto.TextMatrix(i, 2) = Format(rsLlenarConfMont!Riesgo1_A, "#,#00.00")
            RA2 = rsLlenarConfMont!nRA2
            feConfMonto.TextMatrix(i, 3) = Format(rsLlenarConfMont!Riesgo2_A, "#,#00.00")
            NoAceptable = rsLlenarConfMont!nNA
            RNA1 = rsLlenarConfMont!nRN1
            feConfMonto.TextMatrix(i, 4) = Format(rsLlenarConfMont!Riesgo1_N, "#,#00.00")
            RNA2 = rsLlenarConfMont!nRN2
            feConfMonto.TextMatrix(i, 5) = Format(rsLlenarConfMont!Riesgo2_N, "#,#00.00")
            feConfMonto.TextMatrix(i, 6) = rsLlenarConfMont!nNivRiesgo
        rsLlenarConfMont.MoveNext
    Next i
    
RSClose rsLlenarConfMont

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 86 And Shift = 2 Then
        KeyCode = 10
    End If
    
    If KeyCode = 113 And Shift = 0 Then
        KeyCode = 10
    End If
End Sub


