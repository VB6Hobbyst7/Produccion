VERSION 5.00
Begin VB.Form frmCredGastos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gastos del Credito"
   ClientHeight    =   4770
   ClientLeft      =   1650
   ClientTop       =   2565
   ClientWidth     =   9135
   Icon            =   "frmCredGastos.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
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
      Height          =   450
      Left            =   7575
      TabIndex        =   2
      Top             =   4230
      Width           =   1410
   End
   Begin SICMACT.FlexEdit FEGastos 
      Height          =   3450
      Left            =   105
      TabIndex        =   1
      Top             =   705
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   6085
      Cols0           =   5
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "-Aplicado-Nro-Gasto-Monto"
      EncabezadosAnchos=   "400-1200-400-5000-1200"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-C-L-R"
      FormatosEdit    =   "0-0-0-0-2"
      SelectionMode   =   1
      lbUltimaInstancia=   -1  'True
      lbPuntero       =   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Label Label1 
      Caption         =   "Gastos Aplicado al Credito"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   2970
      TabIndex        =   0
      Top             =   195
      Width           =   3195
   End
End
Attribute VB_Name = "frmCredGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub MuestraGastos(ByVal MatGastos As Variant, ByVal pnNumGastos As Integer, _
                        Optional ByVal pbMatrizFormateada As Boolean = False)
Dim i As Integer
Dim k As Integer

    On Error GoTo ErrorMuestraGastos
    
    'Llena Flex de Gastos
    LimpiaFlex FEGastos
    '**CAPI 20080104 **********************
    Dim nGastoAct As Integer
    nGastoAct = 0
    '**************************************
    If pbMatrizFormateada Then
        For i = 0 To UBound(MatGastos) - 1
            '**Modificado por CAPI 20080104
            'If Trim(Right(MatGastos(i, 0), 2)) = "0" Or (Trim(Right(MatGastos(i, 0), 2)) = "1" And (MatGastos(i, 1) = "1" Or MatGastos(i, 1) = "*")) Then
            If Trim(Right(MatGastos(i, 0), 2)) = "0" Or (Trim(Right(MatGastos(i, 0), 2)) = "1" And (CInt(Trim(Right(MatGastos(i, 2), 8))) <> nGastoAct Or MatGastos(i, 1) = "*")) Then
                If CDbl(MatGastos(i, 3)) > 0 Then  'ARCV 14-02-2007
                    k = k + 1
                    FEGastos.AdicionaFila
                    FEGastos.TextMatrix(k, 1) = MatGastos(i, 0) 'Aplicado
                    If Trim(Right(MatGastos(i, 0), 2)) = "1" And MatGastos(i, 1) = "1" Then
                        FEGastos.TextMatrix(k, 2) = "V"
                    Else
                        FEGastos.TextMatrix(k, 2) = MatGastos(i, 1) 'Numero Cuota
                    End If
                    FEGastos.TextMatrix(k, 3) = MatGastos(i, 2) 'Gasto
                    FEGastos.TextMatrix(k, 4) = MatGastos(i, 3) 'Monto
                    '**CAPI 20080104******************************
                    nGastoAct = CInt(Trim(Right(MatGastos(i, 2), 8)))
                    '*********************************************
                End If
            End If
        Next i
    Else
        k = 0
        For i = 0 To pnNumGastos - 1
            If (Trim(Right(MatGastos(i, 0), 2)) = "1" And CInt(MatGastos(i, 1)) = 1) Or (Trim(Right(MatGastos(i, 0), 2)) = "0") Then
                If CDbl(MatGastos(i, 3)) > 0 Then 'ARCV 14-02-2007
                    FEGastos.AdicionaFila
                    k = k + 1
                    FEGastos.TextMatrix(k, 1) = MatGastos(i, 0) 'Aplicado
                    If Trim(Right(MatGastos(i, 0), 2)) = "1" Then
                        FEGastos.TextMatrix(k, 2) = "*"
                    Else
                        FEGastos.TextMatrix(k, 2) = MatGastos(i, 1) 'Numero
                    End If
                    FEGastos.TextMatrix(k, 3) = MatGastos(i, 2) 'Gasto
                    FEGastos.TextMatrix(k, 4) = MatGastos(i, 3) 'Monto
                    FEGastos.RowHeight(k) = 300
                End If
            End If
        Next i
    End If
    
    Me.Show 1
    Exit Sub

ErrorMuestraGastos:
    MsgBox err.Description, vbCritical, "Aviso"

End Sub

'ARCV 14-11-2006
'Public Sub MuestraGastos(ByVal MatGastos As Variant, ByVal pnNumGastos As Integer, _
'                        Optional ByVal pbMatrizFormateada As Boolean = False)
'Dim i As Integer
'Dim k As Integer
'
'    On Error GoTo ErrorMuestraGastos
'
'    'Llena Flex de Gastos
'    LimpiaFlex FEGastos
'
'    If pbMatrizFormateada Then
'        For i = 0 To UBound(MatGastos) - 1
'            FEGastos.AdicionaFila
'            FEGastos.TextMatrix(i + 1, 1) = MatGastos(i, 0)
'            FEGastos.TextMatrix(i + 1, 2) = MatGastos(i, 1)
'            FEGastos.TextMatrix(i + 1, 3) = MatGastos(i, 2)
'            FEGastos.TextMatrix(i + 1, 4) = MatGastos(i, 3)
'        Next i
'    Else
'        k = 0
'        For i = 0 To pnNumGastos - 1
'            If (Trim(Right(MatGastos(i, 0), 2)) = "1" And CInt(MatGastos(i, 1)) = 1) Or (Trim(Right(MatGastos(i, 0), 2)) = "0") Then
'                FEGastos.AdicionaFila
'                k = k + 1
'                FEGastos.TextMatrix(k, 1) = MatGastos(i, 0) 'Aplicado
'                If Trim(Right(MatGastos(i, 0), 2)) = "1" Then
'                    FEGastos.TextMatrix(k, 2) = "*"
'                Else
'                    FEGastos.TextMatrix(k, 2) = MatGastos(i, 1) 'Numero
'                End If
'                FEGastos.TextMatrix(k, 3) = MatGastos(i, 2) 'Gasto
'                FEGastos.TextMatrix(k, 4) = MatGastos(i, 3) 'Monto
'                FEGastos.RowHeight(k) = 300
'            End If
'        Next i
'    End If
'
'    Me.Show 1
'    Exit Sub
'
'ErrorMuestraGastos:
'    MsgBox Err.Description, vbCritical, "Aviso"
'
'End Sub



Private Sub cmdSalir_Click()
    Unload Me
End Sub

