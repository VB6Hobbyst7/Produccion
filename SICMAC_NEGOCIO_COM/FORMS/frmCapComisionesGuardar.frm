VERSION 5.00
Begin VB.Form frmCapComisionesGuardar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Guardar"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5190
   Icon            =   "frmCapComisionesGuardar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnGuardar 
      Caption         =   "Guardar"
      Height          =   300
      Left            =   3240
      TabIndex        =   8
      Top             =   1395
      Width           =   870
   End
   Begin VB.CommandButton btnCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   300
      Left            =   4185
      TabIndex        =   7
      Top             =   1395
      Width           =   870
   End
   Begin VB.Frame Frame1 
      Caption         =   "Versión"
      Height          =   1230
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4965
      Begin VB.TextBox txtGlosa 
         Height          =   300
         Left            =   990
         TabIndex        =   6
         Top             =   765
         Width           =   3615
      End
      Begin VB.TextBox txtFechaRegistro 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3375
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   405
         Width           =   1230
      End
      Begin VB.TextBox txtVersion 
         Enabled         =   0   'False
         Height          =   300
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   405
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "Glosa:"
         Height          =   240
         Left            =   315
         TabIndex        =   5
         Top             =   765
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Registro"
         Height          =   240
         Left            =   2115
         TabIndex        =   3
         Top             =   405
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Versión:"
         Height          =   240
         Left            =   315
         TabIndex        =   1
         Top             =   405
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmCapComisionesGuardar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private objComision As tComision
Private bResp As Boolean

Private Sub btnCancelar_Click()
    Unload Me
End Sub
Public Property Get Comision() As tComision
    Comision = objComision
End Property
Public Property Let Comision(vNewValue As tComision)
    objComision = vNewValue
End Property
Public Function sGosa() As String
    sGosa = Trim(txtGlosa.Text)
End Function
Public Function bRespuesta() As Boolean
    bRespuesta = bResp
End Function
Private Sub btnGuardar_Click()
    bResp = True
    Unload Me
End Sub
Private Sub Form_Initialize()
    bResp = False
End Sub
Private Sub Form_Load()
    txtFechaRegistro.Text = Comision.FechaRegistro
    txtVersion = "V" & IIf(Comision.Version < 10, "0" & CStr(Comision.Version), Comision.Version)
End Sub

