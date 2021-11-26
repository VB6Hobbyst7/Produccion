VERSION 5.00
Begin VB.Form frmChequeDetApert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Apertura"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
   Icon            =   "frmChequeDetApert.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Intervinientes en la Cuenta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2820
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   5655
      Begin SICMACT.FlexEdit feDetalle 
         Height          =   2490
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   4392
         Cols0           =   3
         HighLight       =   2
         AllowUserResizing=   3
         EncabezadosNombres=   "N°-Código-Nombre"
         EncabezadosAnchos=   "350-1500-3400"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-1-0"
         EncabezadosAlineacion=   "C-L-L"
         FormatosEdit    =   "0-0-0"
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
      End
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Height          =   300
      Left            =   60
      TabIndex        =   0
      Top             =   2835
      Width           =   885
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "&Quitar"
      Height          =   300
      Left            =   960
      TabIndex        =   1
      Top             =   2835
      Width           =   885
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   300
      Left            =   4830
      TabIndex        =   2
      Top             =   2835
      Width           =   885
   End
End
Attribute VB_Name = "frmChequeDetApert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************************
'** Nombre : frmChequeDetAper
'** Descripción : Registro de Personas que podran aperturar cuentas segun TI-ERS126-2013
'** Creación : EJVG, 20131217 09:57:00 AM
'***************************************************************************************
Option Explicit
Dim fbAceptar As Boolean
Dim fsListaPersCod As Variant
Dim fbReadOnly As Boolean

Private Sub Form_Activate()
    If fbReadOnly Then
        cmdAgregar.Enabled = False
        cmdQuitar.Enabled = False
        cmdAceptar.Enabled = False
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    fbAceptar = False
End Sub
Private Sub CmdAceptar_Click()
    If Not validaDatosAceptar Then Exit Sub
    fbAceptar = True
    Hide
End Sub
Private Sub cmdAgregar_Click()
    feDetalle.AdicionaFila
    If feDetalle.Visible And feDetalle.Enabled Then feDetalle.SetFocus
    SendKeys "{Enter}"
End Sub
Private Sub cmdQuitar_Click()
    feDetalle.EliminaFila (feDetalle.row)
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub feDetalle_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    If psDataCod = gsCodPersUser Then
        MsgBox "No se puede registrar un Cheque de si mismo", vbInformation, "Aviso"
        feDetalle.EliminaFila pnRow
        Exit Sub
    End If
End Sub
Public Function Inicio(ByVal psListaPersCod As String, Optional ByVal pbReadOnly As Boolean = False) As String
    fbAceptar = False
    fsListaPersCod = psListaPersCod
    Call SetDatosFlex(psListaPersCod)
    fbReadOnly = pbReadOnly
    Show 1
    Inicio = GetDatosFlex()
End Function
Private Sub SetDatosFlex(ByVal psListaPersCod As String)
    Dim MatDatos() As String
    Dim i As Integer
    Call LimpiaFlex(feDetalle)
    If psListaPersCod <> "" Then
        MatDatos = Split(psListaPersCod, ",")
        feDetalle.TabIndex = 0
        For i = 0 To UBound(MatDatos)
            feDetalle.AdicionaFila
            feDetalle.TextMatrix(feDetalle.row, 1) = MatDatos(i)
            SendKeys "{Enter}"
            SendKeys "{Enter}"
        Next
    End If
End Sub
Private Function GetDatosFlex() As String
    Dim lista As String
    Dim i As Integer
    If fbAceptar Then
        For i = 1 To feDetalle.Rows - 1
            lista = lista & feDetalle.TextMatrix(i, 1) & ","
        Next
        lista = Left(lista, Len(lista) - 1)
    Else
        lista = fsListaPersCod
    End If
    GetDatosFlex = lista
End Function
Private Function validaDatosAceptar() As Boolean
    Dim i As Integer
    Dim J As Integer
    validaDatosAceptar = True
    If feDetalle.TextMatrix(1, 0) = "" Then
        MsgBox "Ud. debe de especificar las Personas para la Apertura", vbInformation, "Aviso"
        If feDetalle.Visible And feDetalle.Enabled Then feDetalle.SetFocus
        validaDatosAceptar = False
        Exit Function
    End If
    For i = 1 To feDetalle.Rows - 1
        For J = 1 To feDetalle.Cols - 1
            If feDetalle.ColWidth(J) <> 0 Then
                If Len(Trim(feDetalle.TextMatrix(i, J))) = 0 Then
                    MsgBox "El campo " & UCase(feDetalle.TextMatrix(0, J)) & " está vacio", vbInformation, "Aviso"
                    If feDetalle.Visible And feDetalle.Enabled Then feDetalle.SetFocus
                    feDetalle.TopRow = i
                    feDetalle.row = i
                    feDetalle.Col = J
                    validaDatosAceptar = False
                    Exit Function
                End If
            End If
        Next
    Next
End Function



