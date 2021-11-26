VERSION 5.00
Begin VB.Form frmListarCuentas 
   Caption         =   "Listar Cuentas"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8805
   Icon            =   "frmListarCuentas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Cuentas Contables"
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   7440
         TabIndex        =   4
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "-"
         Height          =   360
         Left            =   7920
         TabIndex        =   3
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "+"
         Height          =   360
         Left            =   7920
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
      Begin Sicmact.FlexEdit fg 
         Height          =   4515
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   7964
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Código-Descripción-Entidad Finc."
         EncabezadosAnchos=   "385-1700-3800-1500"
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
         ColumnasAEditar =   "X-1-X-3"
         TextStyleFixed  =   3
         ListaControles  =   "0-1-0-3"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L"
         FormatosEdit    =   "0-0-0-0"
         CantEntero      =   15
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   390
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmListarCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim oCta As DCtaCont

Public Sub Inicio(ByRef pMatCtaCod As Variant)
    Dim i, k As Integer
    ReDim pMatCtaCod((fg.Rows - 2), 3)
    Do While i < fg.Rows - 1
        pMatCtaCod(i, 0) = fg.TextMatrix(i + 1, 1)
        pMatCtaCod(i, 1) = fg.TextMatrix(i + 1, 2)
        pMatCtaCod(i, 2) = Trim(Left(fg.TextMatrix(i + 1, 3), 20))
        pMatCtaCod(i, 3) = Trim(Right(fg.TextMatrix(i + 1, 3), 20))
        i = i + 1
    Loop
End Sub

Private Sub CargarEntidadesFinc()
    Dim R As ADODB.Recordset
    Set oCta = New DCtaCont
    Set R = oCta.DarEntidadesFinac
    fg.CargaCombo R
    Set oCta = Nothing
    Set R = Nothing
End Sub

Private Sub Form_Load()
    CentraForm Me
    Set oCta = New DCtaCont
    Set rs = oCta.DarCtaContAdeudados
    fg.TipoBusqueda = BuscaGrid
    fg.lbUltimaInstancia = True
    fg.AutoAdd = True
    fg.rsTextBuscar = rs
    CargarEntidadesFinc
    Set oCta = Nothing
    Set rs = Nothing
End Sub

Private Sub cmdNuevo_Click()
    fg.AdicionaFila , Val(fg.TextMatrix(fg.Rows - 1, 0)) + 1
    fg.SetFocus
End Sub

Private Sub cmdEliminar_Click()
    If fg.TextMatrix(fg.Row, 0) <> "" Then
        EliminaCuenta fg.TextMatrix(fg.Row, 1), fg.TextMatrix(fg.Row, 0)
        If fg.TextMatrix(1, 0) = "" Then
            fg.TextMatrix(1, 0) = "1"
        End If
        If fg.Enabled Then
            fg.SetFocus
        End If
    End If
End Sub

Private Sub EliminaCuenta(sCod As String, nItem As Integer)
    fg.EliminaFila fg.Row, False
    EliminaFgObj nItem
    If Len(fg.TextMatrix(1, 1)) > 0 Then
       RefrescaFgObj fg.TextMatrix(fg.Row, 0)
    End If
End Sub

Private Sub EliminaFgObj(nItem As Integer)
    Dim k As Integer
    k = 1
    Do While k < fg.Rows
        If Len(fg.TextMatrix(k, 1)) > 0 Then
            If Val(fg.TextMatrix(k, 0)) = nItem Then
                fg.EliminaFila k, False
            Else
                k = k + 1
            End If
        Else
            k = k + 1
        End If
    Loop
End Sub

Private Sub RefrescaFgObj(nItem As Integer)
    Dim k  As Integer
    For k = 1 To fg.Rows - 1
        If Len(fg.TextMatrix(k, 1)) Then
            If fg.TextMatrix(k, 0) = nItem Then
                fg.RowHeight(k) = 285
            Else
                fg.RowHeight(k) = 0
            End If
        End If
    Next
End Sub

Private Sub Command1_Click()
    If fg.TextMatrix(1, 1) <> "" Then
        Dim ArrayCtaCont() As String
        Inicio ArrayCtaCont
        frmReportes.RecibirArray (ArrayCtaCont)
        Unload Me
    Else
        MsgBox "Debe agregar las Ctas Contables", vbCritical
    End If
End Sub
