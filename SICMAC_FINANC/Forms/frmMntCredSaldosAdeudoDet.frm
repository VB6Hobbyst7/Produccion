VERSION 5.00
Begin VB.Form frmMntCredSaldosAdeudoDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccion de Adeudados para Linea de Creditos"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5985
   Icon            =   "frmMntCredSaldosAdeudoDet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3690
      Left            =   75
      TabIndex        =   2
      Top             =   0
      Width           =   5760
      Begin Sicmact.FlexEdit FeAdeuDet 
         Height          =   2970
         Left            =   150
         TabIndex        =   4
         Top             =   525
         Width           =   5430
         _ExtentX        =   9578
         _ExtentY        =   5239
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Cod. Adeudado-Saldo-Justificacion-Activado"
         EncabezadosAnchos=   "400-1300-1400-1000-1000"
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
         ColumnasAEditar =   "X-X-X-3-4"
         ListaControles  =   "0-0-0-0-4"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-L-L"
         FormatosEdit    =   "0-0-2-0-1"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.CheckBox chkTodos 
         Caption         =   "Todos"
         Height          =   240
         Left            =   4560
         TabIndex        =   3
         Top             =   150
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   3660
      TabIndex        =   1
      Top             =   3765
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   4770
      TabIndex        =   0
      Top             =   3765
      Width           =   1020
   End
End
Attribute VB_Name = "frmMntCredSaldosAdeudoDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public nSaldoTotal As Double
Public cCodPaq As String
Dim nTipoCambio As Currency
Public nSaldoSoles As Double
Public nSaldoDolares As Double
Public bCancel As Boolean

Public Sub Inicio(ByVal psLineaCred As String, ByVal psCodPaq As String)

Dim oAdeu As DCajaCtasIF
'Dim sPersCod As String
Set oAdeu = New DCajaCtasIF
Dim rs As ADODB.Recordset

bCancel = False


'sPersCod = oAdeu.GetPersona_LineaCredito(psLineaCred)
cCodPaq = psCodPaq
FeAdeuDet.rsFlex = oAdeu.GetCredSaldosAdeudoDetalle(psLineaCred, cCodPaq)

FeAdeuDet.ColWidth(5) = 0
Me.Show 1
Set oAdeu = Nothing
End Sub

Private Sub chktodos_Click()
Dim I As Integer

With FeAdeuDet
    If chkTodos.value = 1 Then
        For I = 1 To .Rows - 1
            .Row = I
            .SeleccionaChekTecla
        Next I
    Else
        For I = 1 To .Rows - 1
            .TextMatrix(I, 4) = ""
        Next I
    End If
End With
End Sub

Private Sub cmdAceptar_Click()
Dim I As Integer
Dim oAdeu As DCajaCtasIF

With FeAdeuDet
    'nSaldoTotal = 0
    nSaldoSoles = 0
    nSaldoDolares = 0
    Set oAdeu = New DCajaCtasIF
    Call oAdeu.EliminaCredSaldosAdeudoDetalle(cCodPaq)
    For I = 1 To .Rows - 1
        If .TextMatrix(I, 4) = "." Then
            Call oAdeu.InsertaCredSaldosAdeudoDetalle(cCodPaq, .TextMatrix(I, 1), .TextMatrix(I, 3))
            'ARCV 10-07-2007
            If Mid(Trim(.TextMatrix(I, 1)), 3, 1) = "1" Then
                nSaldoSoles = nSaldoSoles + CDbl(.TextMatrix(I, 2))
            Else
                nSaldoDolares = nSaldoDolares + CDbl(.TextMatrix(I, 2))
            End If
            '-----------
        End If
    Next
    Set oAdeu = Nothing
    bCancel = False
    Unload Me
End With
End Sub

Private Sub cmdCancelar_Click()
    bCancel = True
    Unload Me
End Sub

Private Sub Form_Load()
    Call CentraForm(Me)

Dim oCambio As nTipoCambio
Set oCambio = New nTipoCambio
nTipoCambio = oCambio.EmiteTipoCambio(gdFecSis, TCFijoMes)
Set oCambio = Nothing

End Sub


