VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRHPeriodosNoLaboralesProc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualiza periodos no laborales"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "frmRHPeriodosNoLaboralesProc.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDias 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   5805
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   345
         Left            =   2880
         TabIndex        =   5
         Top             =   750
         Width           =   1170
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
         Height          =   345
         Left            =   1650
         TabIndex        =   4
         Top             =   750
         Width           =   1170
      End
      Begin MSMask.MaskEdBox mskIni 
         Height          =   300
         Left            =   1365
         TabIndex        =   2
         Top             =   270
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFin 
         Height          =   300
         Left            =   3270
         TabIndex        =   3
         Top             =   270
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblAl 
         Caption         =   "Al :"
         Height          =   255
         Left            =   2910
         TabIndex        =   7
         Top             =   300
         Width           =   480
      End
      Begin VB.Label lblDel 
         Caption         =   "Del :"
         Height          =   210
         Left            =   930
         TabIndex        =   6
         Top             =   315
         Width           =   465
      End
   End
   Begin MSComctlLib.ProgressBar P 
      Height          =   285
      Left            =   45
      TabIndex        =   0
      Top             =   1275
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmRHPeriodosNoLaboralesProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdProcesar_Click()
    Dim lnI As Integer
    Dim lnTope As Integer
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim Sql As String
    Dim lsUltMov As String
    
    If Not IsDate(Me.mskIni.Text) Then
        MsgBox "Fecha Inicial no valida.", vbInformation, "Aviso"
        Me.mskIni.SetFocus
        Exit Sub
    ElseIf Not IsDate(Me.mskFin.Text) Then
        MsgBox "Fecha Final no valida.", vbInformation, "Aviso"
        Me.mskFin.SetFocus
        Exit Sub
    ElseIf CDate(Me.mskFin.Text) < CDate(Me.mskIni.Text) Then
        MsgBox "Fecha Final debe ser mayor a la fecha inicial.", vbInformation, "Aviso"
        Me.mskFin.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Desea procesar el recalculo de periodos no laborados de " & Me.mskIni.Text & " a " & Me.mskFin.Text & "  ", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    lnTope = Abs(DateDiff("d", CDate(Me.mskFin.Text), CDate(Me.mskIni.Text)))
    
    Me.P.value = 0
    Me.P.Max = lnTope
    
    oCon.AbreConexion
    
    lsUltMov = GetMovNro(gsCodUser, gsCodAge)
    For lnI = 0 To lnTope
        
        Sql = " Execute dbo.spRHAsistenciaFecha '" & Format(DateAdd("d", lnI, CDate(Me.mskIni.Text)), gcFormatoFecha) & "','" & lsUltMov & "' "
        oCon.Ejecutar Sql
        Me.P.value = lnI
    Next lnI
    
    MsgBox "Proceso terminado.", vbInformation, "Aviso"
    Me.cmdSalir.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub mskFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdProcesar.SetFocus
    End If
End Sub

Private Sub mskIni_GotFocus()
    mskIni.SelStart = 0
    mskIni.SelLength = 50
End Sub

Private Sub mskFin_GotFocus()
    mskFin.SelStart = 0
    mskFin.SelLength = 50
End Sub

Private Sub mskIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.mskFin.SetFocus
    End If
End Sub
