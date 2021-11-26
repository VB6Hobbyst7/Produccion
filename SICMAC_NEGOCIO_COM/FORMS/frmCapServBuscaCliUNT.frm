VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCapServBuscaCliUNT 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   Icon            =   "frmCapServBuscaCliUNT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3420
      TabIndex        =   5
      Top             =   3120
      Width           =   1200
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2100
      TabIndex        =   4
      Top             =   3120
      Width           =   1200
   End
   Begin VB.Frame fraAlumno 
      Caption         =   "Ingrese Nombre del Alumno"
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
      Height          =   735
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   6795
      Begin VB.TextBox txtNomPer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   120
         TabIndex        =   3
         Tag             =   "1"
         Top             =   240
         Width           =   6555
      End
   End
   Begin VB.Frame fraResultado 
      Caption         =   "Resultado Búsqueda"
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
      Height          =   2175
      Left            =   60
      TabIndex        =   0
      Top             =   840
      Width           =   6795
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdAlumno 
         Height          =   1830
         Left            =   135
         TabIndex        =   1
         Top             =   240
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   3228
         _Version        =   393216
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      End
   End
   Begin VB.Label lblTiempo 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1860
      TabIndex        =   6
      Top             =   2325
      Width           =   150
   End
End
Attribute VB_Name = "frmCapServBuscaCliUNT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BuscaPersona()
Dim rsUNT As ADODB.Recordset
Dim clsServ As NCapServicios
'Busqueda por Nombre
Set clsServ = New NCapServicios
Set rsUNT = clsServ.GetUNTAlumnoNombre(Trim(txtNomPer))
Set clsServ = Nothing
grdAlumno.Clear
If rsUNT.EOF And rsUNT.BOF Then
    Me.MousePointer = 0
    Screen.MousePointer = 0
    MsgBox "Datos no Encontrados por favor reintente", vbInformation, Me.Caption
Else
    Set grdAlumno.Recordset = rsUNT
    AjustaGrid
    cmdAceptar.Enabled = True
    grdAlumno.SetFocus
End If
rsUNT.Close
Set rsUNT = Nothing
End Sub

Private Sub CmdAceptar_Click()
Dim nFila As Long
nFila = grdAlumno.Row
frmCapServOpeUNT.mskCodAlumno.Text = Left(grdAlumno.TextMatrix(nFila, 1), 7) & "-" & Right(grdAlumno.TextMatrix(nFila, 1), 2)
frmCapServOpeUNT.lblNombre = grdAlumno.TextMatrix(nFila, 2)
frmCapServOpeUNT.lblEscuela = grdAlumno.TextMatrix(nFila, 3)
frmCapServOpeUNT.lblCodEsc = grdAlumno.TextMatrix(nFila, 4)
frmCapServOpeUNT.lblEscOtros = grdAlumno.TextMatrix(nFila, 4)
frmCapServOpeUNT.sCurr = grdAlumno.TextMatrix(nFila, 5)
frmCapServOpeUNT.fraEstudiante.Enabled = False
frmCapServOpeUNT.fraConcepto.Enabled = True
Unload Me
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = "Busca Alumno Universidad Nacional de Trujillo"
cmdAceptar.Enabled = False
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
End Sub

Private Sub AjustaGrid()
grdAlumno.ColWidth(0) = 150
grdAlumno.ColWidth(1) = 1000
grdAlumno.ColWidth(2) = 3500
grdAlumno.ColWidth(3) = 2000
grdAlumno.ColWidth(4) = 0
grdAlumno.ColWidth(5) = 0
End Sub

Private Sub grdAlumno_KeyPress(KeyAscii As Integer)
If grdAlumno.Rows > 1 And grdAlumno <> "" Then
    If KeyAscii = 13 Then
        CmdAceptar_Click
    End If
End If
End Sub

Private Sub txtNomPer_GotFocus()
txtNomPer.SelStart = 0
txtNomPer.SelLength = Len(txtNomPer.Text)
End Sub

Private Sub txtNomPer_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "%" Or Chr(KeyAscii) = "_" Then
Else
    KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End If
If KeyAscii = 13 Then
    If Len(txtNomPer) <> 0 Then
        BuscaPersona
    Else
        MsgBox "Por favor Ingrese el nombre del alumno", vbInformation, "Aviso"
        cmdAceptar.Enabled = False
        txtNomPer.SetFocus
    End If
End If
End Sub



