VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPreTpo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipo Presupuesto"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   Icon            =   "frmPreTpo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5280
      TabIndex        =   6
      Top             =   1095
      Width           =   1200
   End
   Begin VB.Frame fraPresu 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   45
      TabIndex        =   1
      Top             =   30
      Width           =   6435
      Begin VB.ComboBox cboPresup 
         Height          =   315
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   255
         Width           =   4725
      End
      Begin MSMask.MaskEdBox mskAnio 
         Height          =   300
         Left            =   1590
         TabIndex        =   3
         Top             =   600
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskAnioCopia 
         Height          =   300
         Left            =   4155
         TabIndex        =   7
         Top             =   600
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblAnioCopia 
         Caption         =   "Año Copia (Nuevo):"
         Height          =   195
         Left            =   2490
         TabIndex        =   8
         Top             =   660
         Width           =   1500
      End
      Begin VB.Label lblAnio 
         Caption         =   "Año Base:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   645
         Width           =   990
      End
      Begin VB.Label lblPresupuesto 
         Caption         =   "Presupuesto Base :"
         Height          =   225
         Left            =   105
         TabIndex        =   4
         Top             =   285
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   4035
      TabIndex        =   0
      Top             =   1095
      Width           =   1200
   End
End
Attribute VB_Name = "frmPreTpo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsCodigo As String
Dim lnAnio As Long
Dim lnAnioCopia As Long

Public Sub Ini(psCodigo As String, pnAnio As Long, pnAnioCopia As Long)
    Me.Show 1
    psCodigo = lsCodigo
    pnAnio = lnAnio
    pnAnioCopia = lnAnioCopia
End Sub

Private Sub cmdAceptar_Click()
    If Me.cboPresup.Text = "" Then
        MsgBox "Debe elejir un presupuesto.", vbInformation, "Aviso"
        Me.cboPresup.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(Me.mskAnio.Text) Then
        MsgBox "Debe elejir un Año Valido.", vbInformation, "Aviso"
        Me.mskAnio.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(Me.mskAnioCopia.Text) Then
        MsgBox "Debe elejir un Año Valido.", vbInformation, "Aviso"
        Me.mskAnioCopia.SetFocus
        Exit Sub
    End If
    
    lsCodigo = Trim(Right(Me.cboPresup.Text, 5))
    lnAnio = Me.mskAnio.Text
    lnAnioCopia = Me.mskAnioCopia.Text
    
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    lsCodigo = ""
    lnAnio = -1
    lnAnioCopia = -1
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oPresup As DPresupuesto
    Set oPresup = New DPresupuesto


    Set rs = oPresup.GetPresupuesto(True)
    
    CargaCombo rs, Me.cboPresup, , 1, 0
    
    Me.mskAnio.Text = Format(gdFecSis, "yyyy")
    Me.mskAnioCopia.Text = Format(gdFecSis, "yyyy")
End Sub
