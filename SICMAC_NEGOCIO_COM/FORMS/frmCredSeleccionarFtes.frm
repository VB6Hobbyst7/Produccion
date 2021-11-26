VERSION 5.00
Begin VB.Form frmCredSeleccionarFtes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar Fuentes de Ingreso del Crédito"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6570
   Icon            =   "frmCredSeleccionarFtes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   5025
      TabIndex        =   2
      Top             =   2775
      Width           =   1440
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   390
      Left            =   3600
      TabIndex        =   1
      Top             =   2775
      Width           =   1440
   End
   Begin SICMACT.FlexEdit FEFuentes 
      Height          =   2640
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   4657
      Cols0           =   8
      HighLight       =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      EncabezadosNombres=   "Nº-OK-cNumFuente-Fuente de Ingreso-cPersCod-dPersFecEval-dPersFecCaduc-nPersFteIngTipo"
      EncabezadosAnchos=   "400-400-0-5500-0-0-0-0"
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
      ColumnasAEditar =   "X-1-X-3-X-X-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-4-0-0-0-0-0-0"
      BackColorControl=   65535
      BackColorControl=   65535
      BackColorControl=   65535
      EncabezadosAlineacion=   "C-L-C-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-5-5-0"
      AvanceCeldas    =   1
      TextArray0      =   "Nº"
      lbEditarFlex    =   -1  'True
      lbFlexDuplicados=   0   'False
      lbFormatoCol    =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483635
   End
End
Attribute VB_Name = "frmCredSeleccionarFtes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MatFteIngresos As Variant

Public Function Inicio(ByVal psPersCod As String) As Variant

Dim oPersona As UPersona_Cli
Dim i As Integer
Dim MatFte As Variant

Dim rsFteIng As ADODB.Recordset
Dim rsFIDep As ADODB.Recordset
Dim rsFIInd As ADODB.Recordset
Dim oCred As COMNCredito.NCOMCredito

    On Error GoTo ErrorCargaFuentesIngreso
    Set oPersona = New UPersona_Cli
    Call oPersona.RecuperaPersona_Solicitud(psPersCod)
        
    Set oCred = New COMNCredito.NCOMCredito
            
    MatFte = oPersona.FiltraFuentesIngresoPorRazonSocial
    
    FEFuentes.Clear
    FEFuentes.Rows = 2
    FEFuentes.FormaCabecera
    
    If IsArray(MatFte) Then
        With FEFuentes
            For i = 0 To UBound(MatFte) - 1
                .AdicionaFila , , True
                .TextMatrix(i + 1, 1) = ""
                .TextMatrix(i + 1, 2) = MatFte(i, 8)
                .TextMatrix(i + 1, 3) = MatFte(i, 2)
                .TextMatrix(i + 1, 4) = MatFte(i, 6)
                Call oCred.CargarFtesIngreso(rsFteIng, rsFIDep, rsFIInd, psPersCod, , i)
                Call oPersona.RecuperaFtesdeIngreso(psPersCod, rsFteIng)
                Call oPersona.RecuperaFtesIngresoDependiente(i, rsFIDep)
                Call oPersona.RecuperaFtesIngresoIndependiente(i, rsFIInd)
                .TextMatrix(i + 1, 5) = oPersona.ObtenerFteIngFecEval(i, IIf(oPersona.ObtenerFteIngIngresoTipo(i) = gPersFteIngresoTipoDependiente, oPersona.ObtenerFteIngIngresoNumeroFteDep(i) - 1, oPersona.ObtenerFteIngIngresoNumeroFteIndep(i) - 1))
                .TextMatrix(i + 1, 6) = oPersona.ObtenerFteIngFecCaducac(i, IIf(oPersona.ObtenerFteIngIngresoTipo(i) = gPersFteIngresoTipoDependiente, oPersona.ObtenerFteIngIngresoNumeroFteDep(i) - 1, oPersona.ObtenerFteIngIngresoNumeroFteIndep(i) - 1))
                .TextMatrix(i + 1, 7) = oPersona.ObtenerFteIngIngresoTipo(i)
            Next i
        End With
        Set oCred = Nothing
    End If
    
    Set oPersona = Nothing
    Me.Show 1


    Inicio = MatFteIngresos

    Exit Function

ErrorCargaFuentesIngreso:
        MsgBox Err.Description, vbCritical, "Aviso"

End Function

Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    'No seleccionar ninguna Fte
    ReDim MatFteIngresos(0, 0)
    Unload Me
End Sub

Private Sub Form_Load()
    Call CentraForm(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
Dim nFilas As Integer
    
    With FEFuentes
        
        nFilas = 0
        For i = 0 To .Rows - 2
            If .TextMatrix(i + 1, 1) = "." Then
                nFilas = nFilas + 1
            End If
        Next i
        
        ReDim MatFteIngresos(nFilas, 6)
        nFilas = 0
        For i = 0 To .Rows - 2
            If .TextMatrix(i + 1, 1) = "." Then
                nFilas = nFilas + 1
                MatFteIngresos(nFilas - 1, 0) = .TextMatrix(i + 1, 2)
                MatFteIngresos(nFilas - 1, 1) = .TextMatrix(i + 1, 3)
                MatFteIngresos(nFilas - 1, 2) = .TextMatrix(i + 1, 4)
                MatFteIngresos(nFilas - 1, 3) = .TextMatrix(i + 1, 5)
                MatFteIngresos(nFilas - 1, 4) = .TextMatrix(i + 1, 6)
                MatFteIngresos(nFilas - 1, 5) = .TextMatrix(i + 1, 7)
            End If
        Next i
    End With
    
End Sub
