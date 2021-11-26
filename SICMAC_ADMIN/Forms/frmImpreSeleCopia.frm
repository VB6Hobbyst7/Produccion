VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmImpreSeleCopia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresión de Copias"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   2850
      TabIndex        =   3
      Top             =   1980
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Imprimir"
      Height          =   345
      Left            =   1590
      TabIndex        =   2
      Top             =   1980
      Width           =   1215
   End
   Begin VB.ListBox lstOpc 
      Appearance      =   0  'Flat
      Height          =   1380
      ItemData        =   "frmImpreSeleCopia.frx":0000
      Left            =   75
      List            =   "frmImpreSeleCopia.frx":0010
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   540
      Width           =   3975
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   270
      Left            =   105
      TabIndex        =   4
      Top             =   2025
      Visible         =   0   'False
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   476
      _Version        =   393217
      Enabled         =   0   'False
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmImpreSeleCopia.frx":0040
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblDoc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ORDEN DE COMPRA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   465
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   3975
   End
End
Attribute VB_Name = "frmImpreSeleCopia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsImpre As String
Dim lsDoc   As String
Dim K       As Integer

Public Sub Inicio(psImpre As String, psDoc As String)
lsImpre = psImpre
lsDoc = psDoc
Me.Show 1
End Sub

Private Sub CmdAceptar_Click()
    Dim oPrevio As clsPrevio
    Set oPrevio = New clsPrevio
    Dim sImpre As String
    sImpre = ""
    For K = 0 To lstOpc.ListCount - 1
       If lstOpc.Selected(K) Then
          If sImpre <> "" Then
             sImpre = sImpre & oImpresora.gPrnSaltoPagina
          End If
          sImpre = sImpre & lsImpre & Space(70) & BON & lstOpc.List(K) & BOFF
       End If
    Next
    oPrevio.Show sImpre, lsDoc, False, 66
    Unload Me
End Sub

Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
    LblDoc.Caption = lsDoc
    For K = 0 To lstOpc.ListCount - 1
       lstOpc.Selected(K) = True
       If UCase(lsDoc) = "ORDEN DE COMPRA" And K = 1 Then
          lstOpc.List(K) = "COMPRA"
       End If
    Next
End Sub

