VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLogProSelRptProcesos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consultar Procesos"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "frmLogProSelRptProcesos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Reporte"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.ComboBox cboProceso 
         Height          =   315
         ItemData        =   "frmLogProSelRptProcesos.frx":08CA
         Left            =   1080
         List            =   "frmLogProSelRptProcesos.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   4575
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Left            =   4440
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Top             =   1440
         Width           =   1275
      End
      Begin MSMask.MaskEdBox txtmesIni 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   840
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtmesFin 
         Height          =   315
         Left            =   3120
         TabIndex        =   4
         Top             =   840
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtaño 
         Height          =   315
         Left            =   4800
         TabIndex        =   7
         Top             =   840
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Proceso"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         Height          =   195
         Left            =   4440
         TabIndex        =   8
         Top             =   960
         Width           =   285
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes Término"
         Height          =   195
         Left            =   2160
         TabIndex        =   6
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Mes Inicio"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmLogProSelRptProcesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImprimir_Click()
    Dim cAbre As String, nPos As Integer
    If txtaño.Text = "____" Then Exit Sub
    If txtmesIni.Text = "__" Then Exit Sub
    If txtmesFin.Text = "__" Then Exit Sub
    nPos = InStr(1, cboProceso.Text, " - ")
    If nPos > 0 Then cAbre = Mid(cboProceso.Text, nPos + 3)
    ImprimeListaProceso txtaño.Text, cAbre, txtmesIni, txtmesFin
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CentraForm Me
    txtaño.Text = Year(gdFecSis)
    txtmesIni.Text = "01"
    txtmesFin.Text = Format(Month(gdFecSis), "00")
    CargarCombo
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frmLogProSelRptProcesos = Nothing
End Sub

Private Sub txtaño_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdImprimir.SetFocus
End Sub

Private Sub txtmesFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtaño.SetFocus
End Sub

Private Sub txtmesIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtmesFin.SetFocus
End Sub

Private Sub CargarCombo()
On Error GoTo CargarComboErr
    Dim Rs As ADODB.Recordset
    Set Rs = CargarTipos
    cboProceso.Clear
    Do While Not Rs.EOF
        cboProceso.AddItem Rs!cProSelTpoDescripcion & " - " & Rs!cAbreviatura
        Rs.MoveNext
    Loop
    If cboProceso.ListCount > 0 Then cboProceso.ListIndex = 0
    Exit Sub
CargarComboErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Sub
