VERSION 5.00
Begin VB.Form frmCaracImpresion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Caracteres de Impresion"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   Icon            =   "frmCaracImpresion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   420
      Left            =   2655
      TabIndex        =   3
      Top             =   735
      Width           =   1515
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   420
      Left            =   960
      TabIndex        =   2
      Top             =   735
      Width           =   1515
   End
   Begin VB.ComboBox CmbImpresora 
      Height          =   315
      Left            =   1815
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   210
      Width           =   2880
   End
   Begin VB.Label Label1 
      Caption         =   "Marca de Impresora :"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1560
   End
End
Attribute VB_Name = "frmCaracImpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAceptar_Click()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia Right(CmbImpresora.Text, 1)
    
    gImpresora = Right(CmbImpresora.Text, 1)
    
    oImp.SetImpreSetup oImp.GetMaquina, Right(CmbImpresora.Text, 1)
    
    MsgBox "Caracteres Configurados", vbInformation, "Aviso"
    
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
     
    CmbImpresora.AddItem "EPSON" & Space(100) & gEPSON
    CmbImpresora.AddItem "HEWLETT PACKARD" & Space(100) & gHEWLETT_PACKARD
    CmbImpresora.AddItem "IBM" & Space(100) & gIBM
 
    CmbImpresora.ListIndex = oImp.GetImpreSetup(oImp.GetMaquina) - 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CierraConexion
End Sub
