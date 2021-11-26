VERSION 5.00
Begin VB.Form frmRHCierreDia 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   Icon            =   "frmRHCierreDia.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4950
      TabIndex        =   1
      Top             =   1815
      Width           =   975
   End
   Begin VB.Frame fraCierre 
      Caption         =   "Cierre de Dia de Recursos Humanos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1650
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   5850
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "&Generar Cierre de Día"
         Height          =   375
         Left            =   1260
         TabIndex        =   3
         Top             =   1185
         Width           =   3360
      End
      Begin VB.Label lblCierre 
         Caption         =   $"frmRHCierreDia.frx":030A
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   840
         Left            =   75
         TabIndex        =   2
         Top             =   255
         Width           =   5700
      End
   End
End
Attribute VB_Name = "frmRHCierreDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGenerar_Click()
    Dim oPla As NRHProcesosCierre
    Dim oPrevio As Previo.clsPrevio
    Set oPla = New NRHProcesosCierre
    Set oPrevio = New Previo.clsPrevio
    Dim lsCadena As String
    
    lsCadena = oPla.CierreDia(gdFecSis, gsNomAge, gsEmpresa, gdFecSis)
    
    If lsCadena <> "" Then oPrevio.Show lsCadena, Caption, True, 66
    
    Set oPla = Nothing
    Set oPrevio = Nothing
    Me.cmdGenerar.Enabled = False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Sub Ini(psCaption As String)
    Caption = psCaption
    Me.Show 1
End Sub

