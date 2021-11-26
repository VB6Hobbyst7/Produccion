VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   1395
   ClientTop       =   2055
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   7200
   Begin RichTextLib.RichTextBox R 
      Height          =   210
      Left            =   1095
      TabIndex        =   3
      Top             =   4020
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   370
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command2"
      Height          =   435
      Left            =   3465
      TabIndex        =   2
      Top             =   2190
      Width           =   2610
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   435
      Left            =   3375
      TabIndex        =   1
      Top             =   1410
      Width           =   2610
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   810
      TabIndex        =   0
      Top             =   3075
      Width           =   2085
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim previ As Previo.clsPrevio
Private Sub Command1_Click()
    Set previ = New Previo.clsPrevio
    Me.R.LoadFile App.Path & "\AAA.txt", 1
    previ.Show R.Text, "eJEMPLO", True
End Sub

Private Sub Command2_Click()
    Set previ = New Previo.clsPrevio
    previ.ShowImpreSpool "Ejemplo de Impresion"
End Sub

Private Sub Command3_Click()
    Set previ = New Previo.clsPrevio
    previ.PrintSpool "LPT1", "EJEMPLO DE IMPRESION SPOOL"
End Sub
