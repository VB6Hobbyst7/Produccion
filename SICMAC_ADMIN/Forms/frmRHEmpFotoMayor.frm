VERSION 5.00
Begin VB.Form frmRHEmpFotoMayor 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Foto : "
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   Icon            =   "frmRHEmpFotoMayor.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   330
      Left            =   2220
      TabIndex        =   0
      Top             =   6080
      Width           =   1095
   End
   Begin VB.Image picFoto 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   6015
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmRHEmpFotoMayor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Ini(psNombre As String, pic As Image)
    Set picFoto.Picture = pic.Picture
    
    Caption = "Foto :" & psNombre
    
    picFoto.Stretch = True
    
    Show 1
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
