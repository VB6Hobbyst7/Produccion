VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCredPendPago 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10350
   Icon            =   "frmCredPendPago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   4845
      Width           =   1215
   End
   Begin VB.Frame fraCreditos 
      Caption         =   "Créditos Pendientes de Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   4740
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   10275
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCred 
         Height          =   3735
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   10080
         _ExtentX        =   17780
         _ExtentY        =   6588
         _Version        =   393216
         BackColor       =   -2147483624
         SelectionMode   =   1
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
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   0
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   4140
         Width           =   10080
      End
   End
End
Attribute VB_Name = "frmCredPendPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SetupGrid()
grdCred.Clear
grdCred.Rows = 2
grdCred.Cols = 11
grdCred.ColWidth(0) = 250
grdCred.ColWidth(1) = 2000
grdCred.ColWidth(2) = 300
grdCred.ColWidth(3) = 1200
grdCred.ColWidth(4) = 400
grdCred.ColWidth(5) = 400
grdCred.ColWidth(6) = 1000
grdCred.ColWidth(7) = 1000
grdCred.ColWidth(8) = 1000
grdCred.ColWidth(9) = 1000
grdCred.ColWidth(10) = 1000

grdCred.TextMatrix(0, 0) = "#"
grdCred.TextMatrix(0, 1) = "Cuenta"
grdCred.TextMatrix(0, 2) = ""
grdCred.TextMatrix(0, 3) = "Vencimiento"
grdCred.TextMatrix(0, 4) = "Nro"
grdCred.TextMatrix(0, 5) = ""
grdCred.TextMatrix(0, 6) = "Cuota"
grdCred.TextMatrix(0, 7) = "Mora"
grdCred.TextMatrix(0, 8) = "Gasto"
grdCred.TextMatrix(0, 9) = "Por Pagar"
grdCred.TextMatrix(0, 10) = "Min. Pagar"

grdCred.ColAlignment(0) = 4
grdCred.ColAlignment(1) = 4
grdCred.ColAlignment(2) = 4
grdCred.ColAlignment(3) = 4
grdCred.ColAlignment(4) = 4
grdCred.ColAlignment(5) = 4
grdCred.ColAlignment(6) = 7
grdCred.ColAlignment(7) = 7
grdCred.ColAlignment(8) = 7
grdCred.ColAlignment(9) = 7
grdCred.ColAlignment(10) = 7

grdCred.ColAlignmentFixed(0) = 4
grdCred.ColAlignmentFixed(1) = 4
grdCred.ColAlignmentFixed(2) = 4
grdCred.ColAlignmentFixed(3) = 4
grdCred.ColAlignmentFixed(4) = 4
grdCred.ColAlignmentFixed(5) = 4
grdCred.ColAlignmentFixed(6) = 4
grdCred.ColAlignmentFixed(7) = 4
grdCred.ColAlignmentFixed(8) = 4
grdCred.ColAlignmentFixed(9) = 4
grdCred.ColAlignmentFixed(10) = 4
End Sub

Public Sub Inicia(ByVal rsCred As ADODB.Recordset)
Dim sSimbMoneda As String
Dim nCuentas As Integer, J As Integer
SetupGrid
nCuentas = 0
rsCred.MoveFirst
Do While Not rsCred.EOF
    nCuentas = nCuentas + 1
    If nCuentas > 1 Then grdCred.Rows = grdCred.Rows + 1
    If Mid(rsCred(0), 9, 1) = "2" Then
        For J = 1 To grdCred.Cols - 1
            grdCred.Col = J
            grdCred.CellBackColor = &HC0FFC0
        Next J
        sSimbMoneda = "US$"
    Else
        sSimbMoneda = "S/."
    End If
    grdCred.Row = grdCred.Rows - 1
    grdCred.TextMatrix(grdCred.Row, 0) = Trim(grdCred.Row)
    grdCred.TextMatrix(grdCred.Row, 1) = Trim(rsCred("Cuenta"))
    grdCred.TextMatrix(grdCred.Row, 2) = Trim(rsCred("Estado"))
    grdCred.TextMatrix(grdCred.Row, 3) = IIf(IsNull(rsCred("Vencimiento")), "", Format$(rsCred("Vencimiento"), "dd/mm/yyyy"))
    grdCred.TextMatrix(grdCred.Row, 4) = Trim(rsCred("CuoVenc"))
    grdCred.TextMatrix(grdCred.Row, 5) = Trim(sSimbMoneda)
    grdCred.TextMatrix(grdCred.Row, 6) = Format$(rsCred("Cuota"), "#,##0.00")
    grdCred.TextMatrix(grdCred.Row, 7) = Format$(rsCred("Mora"), "#,##0.00")
    grdCred.TextMatrix(grdCred.Row, 8) = Format$(rsCred("Gastos"), "#,##0.00")
    grdCred.TextMatrix(grdCred.Row, 9) = Format$(rsCred("Cuota") + rsCred("Mora") + rsCred("Gastos"), "#,##0.00")
    grdCred.TextMatrix(grdCred.Row, 10) = Format$(rsCred("MinPagar"), "#,##0.00")
    rsCred.MoveNext
Loop
Me.Show 1
End Sub

Private Sub CmdAceptar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
lblMensaje = "LA CUENTA SE BLOQUEARA AUTOMATICAMENTE. CONSULTE CON SU ADMINISTRADOR PARA SU DESBLOQUEO Y CANCELACION."
End Sub
