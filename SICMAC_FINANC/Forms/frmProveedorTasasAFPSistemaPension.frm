VERSION 5.00
Begin VB.Form frmProveedorTasasAFPSistemaPension 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Porcetanjes Comisiones y Seguro AFP"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6480
   Icon            =   "frmProveedorTasasAFPSistemaPension.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   320
      Left            =   5400
      TabIndex        =   4
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   320
      Left            =   2175
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   320
      Left            =   150
      TabIndex        =   2
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   320
      Left            =   3165
      TabIndex        =   1
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
   End
   Begin Sicmact.FlexEdit fg 
      Height          =   3045
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5371
      Cols0           =   7
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Cod. AFP-AFP-Com. Flujo(%)-Com. Saldo(%)-Prima Seguro(%)-aux"
      EncabezadosAnchos=   "0-0-2000-1300-1300-1450-0"
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
      ColumnasAEditar =   "X-X-X-3-4-5-X"
      ListaControles  =   "0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-L-L-C-C-C-L"
      FormatosEdit    =   "0-0-0-2-2-2-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      lbPuntero       =   -1  'True
      lbOrdenaCol     =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmProveedorTasasAFPSistemaPension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'** Nombre : frmProveedorTasasAFPSistemaPension
'** Descripción : Formulario para mantenimiento de las tasas de AFP
'** Creación : EJVG, 20140804 05:00:00 PM
'******************************************************************
Dim fnRowNoMove As Integer

Private Sub cmdAceptar_Click()
    Dim oPSP As NProveedorSistPens
    Dim fila As Integer
    Dim bExito As Boolean
        
    On Error GoTo ErrAceptar
    
    If fg.TextMatrix(1, 0) = "" Then
        MsgBox "No existen datos a grabar", vbInformation, "Aviso"
        Exit Sub
    End If
    
    fila = fg.row
    If MsgBox("¿Está seguro de actualizar las tasas de " & UCase(fg.TextMatrix(fila, 2)), vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    cmdAceptar.Enabled = False
    Set oPSP = New NProveedorSistPens
    
    bExito = oPSP.ActualizaTasasAFP(fg.TextMatrix(fila, 1), CCur(fg.TextMatrix(fila, 3)), CCur(fg.TextMatrix(fila, 4)), CCur(fg.TextMatrix(fila, 5)))
    
    If bExito Then
        MsgBox "Se ha actualizado satisfactoriamente las Tasas de " & UCase(fg.TextMatrix(fila, 2)), vbInformation, "Aviso"
        HabilitaControles True
        fnRowNoMove = 0
    Else
        MsgBox "Ha sucedido un error, si el problema persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If

    cmdAceptar.Enabled = True
    Screen.MousePointer = 0
    Exit Sub
ErrAceptar:
    cmdAceptar.Enabled = True
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdCancelar_Click()
    HabilitaControles True
    fnRowNoMove = 0
    CargarTasasAFP
End Sub
Private Sub cmdEditar_Click()
    If fg.TextMatrix(1, 0) <> "" Then
        fnRowNoMove = fg.row
        HabilitaControles False
        fg.SetFocus
        SendKeys "{Enter}"
    Else
        MsgBox "No existen datos", vbInformation, "Aviso"
    End If
End Sub
Private Sub cmdNuevo_Click()
    fg.AdicionaFila
    fg.SetFocus
    SendKeys "{Enter}"
    fnRowNoMove = fg.Rows - 1
    HabilitaControles False
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub HabilitaControles(ByVal pbHabilita As Boolean)
    cmdEditar.Visible = pbHabilita
    cmdAceptar.Visible = Not pbHabilita
    cmdCancelar.Visible = Not pbHabilita
    fg.lbEditarFlex = Not pbHabilita
End Sub
Private Sub fg_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim mat() As String
    mat = Split(fg.ColumnasAEditar, "-")
    
    If mat(pnCol) = "X" Then
        MsgBox "Esta columna no es editable", vbInformation, "Aviso"
        Cancel = False
        Exit Sub
    End If
    If pnCol = 3 Or pnCol = 4 Or pnCol = 5 Then
        If Not IsNumeric(fg.TextMatrix(pnRow, pnCol)) Then
            MsgBox "Ingrese un monto mayor a cero", vbInformation, "Aviso"
            Cancel = False
            Exit Sub
        End If
    End If
End Sub
Private Sub fg_RowColChange()
    If fnRowNoMove > 0 Then
        fg.row = fnRowNoMove
    End If
End Sub
Private Sub Form_Load()
    cmdCancelar_Click
End Sub
Private Function CargarTasasAFP() As Boolean
    Dim oPSP As New NProveedorSistPens
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo ErrCargar
    Screen.MousePointer = 11
    FormateaFlex fg
    Set rs = oPSP.ListaTasasAFP
    If Not rs.EOF Then
        Do While Not rs.EOF
            fg.AdicionaFila
            i = fg.row
            fg.TextMatrix(i, 1) = rs!cPersCod
            fg.TextMatrix(i, 2) = rs!cPersNombre
            fg.TextMatrix(i, 3) = Format(rs!nTasaComiFlujo, gsFormatoNumeroView)
            fg.TextMatrix(i, 4) = Format(rs!nTasaComiSaldo, gsFormatoNumeroView)
            fg.TextMatrix(i, 5) = Format(rs!nTasaSeguro, gsFormatoNumeroView)
            rs.MoveNext
        Loop
        CargarTasasAFP = True
    End If
    RSClose rs
    Set oPSP = Nothing
    Screen.MousePointer = 0
    Exit Function
ErrCargar:
    Screen.MousePointer = 0
    CargarTasasAFP = False
    MsgBox Err.Description, vbCritical, "Aviso"
End Function
