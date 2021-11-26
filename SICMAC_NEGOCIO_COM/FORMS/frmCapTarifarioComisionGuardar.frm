VERSION 5.00
Begin VB.Form frmCapTarifarioGuardar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Guardar"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5190
   Icon            =   "frmCapTarifarioComisionGuardar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnGuardar 
      Caption         =   "Guardar"
      Height          =   300
      Left            =   3240
      TabIndex        =   8
      Top             =   1350
      Width           =   870
   End
   Begin VB.CommandButton btnCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   300
      Left            =   4185
      TabIndex        =   7
      Top             =   1350
      Width           =   870
   End
   Begin VB.Frame Frame1 
      Caption         =   "Versión"
      Height          =   1140
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4965
      Begin VB.TextBox txtGlosa 
         Height          =   300
         Left            =   990
         TabIndex        =   6
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox txtFechaRegistro 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3375
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   1230
      End
      Begin VB.TextBox txtVersion 
         Enabled         =   0   'False
         Height          =   300
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "Glosa:"
         Height          =   240
         Left            =   315
         TabIndex        =   5
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Registro"
         Height          =   240
         Left            =   2115
         TabIndex        =   3
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Versión:"
         Height          =   240
         Left            =   315
         TabIndex        =   1
         Top             =   360
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmCapTarifarioGuardar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************************
'* NOMBRE         : frmCapTarifarioGuardar
'* DESCRIPCION    : Proyecto - Tarifario Versionado - Guardar las versiones de tasas y comisiones
'* CREACION       : RIRO, 20160420 10:00 AM
'************************************************************************************************************

Option Explicit

Private bResp As Boolean
Private nTipo_ As Integer ' 1=comision, 2=tasas
Private sGlosa_ As String
Private dFechaRegistro_ As Date
Private nVersion_ As Integer
Private objComision As tComision
'

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
    sGosa = sGlosa_
End Function

Public Property Get dFechaRegistro() As Date
    dFechaRegistro = dFechaRegistro_
End Property
Public Property Let dFechaRegistro(vNewValue As Date)
    dFechaRegistro_ = vNewValue
End Property
Public Property Get nVersion() As Integer
    nVersion = nVersion_
End Property
Public Property Let nVersion(vNewValue As Integer)
    nVersion_ = vNewValue
End Property
Public Property Get nTipo() As Integer
    nTipo = nTipo_
End Property
Public Property Let nTipo(vNewValue As Integer)
    nTipo_ = vNewValue
End Property
Public Function bRespuesta() As Boolean
    bRespuesta = bResp
End Function
Private Sub btnGuardar_Click()
    bResp = True
    sGlosa_ = txtGlosa.Text
    Unload Me
End Sub
Private Sub Form_Initialize()
    bResp = False
End Sub
Private Sub Form_Load()
    If nTipo = 1 Then ' comision
        txtFechaRegistro.Text = Comision.FechaRegistro
        txtVersion = "V" & IIf(Comision.Version < 10, "0" & CStr(Comision.Version), Comision.Version)
        
    ElseIf nTipo = 2 Then 'tasas
        txtFechaRegistro.Text = dFechaRegistro
        txtVersion = "V" & IIf(nVersion < 10, "0" & CStr(nVersion), nVersion)
        
    End If
End Sub
Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        btnGuardar.SetFocus
    End If
End Sub





















