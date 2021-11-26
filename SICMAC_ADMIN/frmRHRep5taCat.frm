VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRHRep5taCat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Certificados de 5ta Categoria"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   Icon            =   "frmRHRep5taCat.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   6375
      TabIndex        =   4
      Top             =   1245
      Width           =   1185
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   345
      Left            =   5160
      TabIndex        =   3
      Top             =   1245
      Width           =   1185
   End
   Begin VB.Frame fra5ta 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Certificado de 5ta Categoria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1140
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   7530
      Begin VB.CheckBox chkIES 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Incluir IES"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1830
         TabIndex        =   7
         Top             =   757
         Width           =   1950
      End
      Begin MSMask.MaskEdBox mskAnio 
         Height          =   315
         Left            =   6300
         TabIndex        =   6
         Top             =   690
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Todos"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   135
         TabIndex        =   5
         Top             =   765
         Width           =   1335
      End
      Begin Sicmact.TxtBuscar txtPersona 
         Height          =   315
         Left            =   135
         TabIndex        =   1
         Top             =   330
         Width           =   1695
         _extentx        =   2990
         _extenty        =   556
         appearance      =   0
         appearance      =   0
         font            =   "frmRHRep5taCat.frx":030A
         appearance      =   0
         tipobusqueda    =   7
      End
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1830
         TabIndex        =   2
         Top             =   330
         Width           =   5640
      End
   End
End
Attribute VB_Name = "frmRHRep5taCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents oPla As NActualizaDatosConPlanilla
Attribute oPla.VB_VarHelpID = -1
Dim Progress As clsProgressBar

Dim lscadena As String
Dim lsDoc As String

Public Function Ini(pForm As Form) As String
    Me.Show 0, pForm
    Ini = lscadena
End Function

Private Sub cmdProcesar_Click()
    Set oPla = New NActualizaDatosConPlanilla
    
    If Me.txtPersona.Text = "" And Me.chkTodos.value = 0 Then
        MsgBox "Debe ingresar un empleado valido.", vbInformation, "Aviso"
        Me.txtPersona.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(Me.mskAnio.Text) Then
        MsgBox "Debe ingresar un año valido.", vbInformation, "Aviso"
        Me.mskAnio.SetFocus
        Exit Sub
    End If
    
    
    If Me.chkTodos.value = 0 Then
        lscadena = oPla.GetRep5taCat(Me.txtPersona.Text, lsDoc, lblNombre.Caption, Me.mskAnio.Text, gsRUC, gdFecSis, Me.chkIES.value)
    Else
        lscadena = oPla.GetRep5taCatTot(Me.mskAnio.Text, gsRUC, gdFecSis, Me.chkIES.value)
    End If
    
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    lscadena = ""
    Unload Me
End Sub

Private Sub Form_Load()
    Set Progress = New clsProgressBar
End Sub

Private Sub mskAnio_GotFocus()
    mskAnio.SelStart = 0
    mskAnio.SelLength = 50
End Sub

Private Sub mskAnio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdProcesar.SetFocus
    End If
End Sub

Private Sub txtPersona_EmiteDatos()
    lblNombre.Caption = txtPersona.psDescripcion
    
    If txtPersona.psDescripcion <> "" Then
        lsDoc = Trim(txtPersona.rsDocPers.Fields(1))
    End If
    
    Me.mskAnio.SetFocus
End Sub

Private Sub oPla_CloseProgress()
    Progress.CloseForm Me
End Sub

Private Sub oPla_Progress(pnValor As Long, pnTotal As Long)
    Progress.Max = pnTotal
    Progress.Progress pnValor, "Generando Reporte"
End Sub

Private Sub oPla_ShowProgress()
    Progress.ShowForm Me
End Sub
