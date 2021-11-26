VERSION 5.00
Begin VB.Form frmUsuarioLeasing 
   Caption         =   "Actualizar Usuario"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7560
   Icon            =   "frmUsuarioLeasing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   7560
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   3240
      TabIndex        =   13
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4680
      TabIndex        =   12
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "SAF"
      Height          =   1095
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   7215
      Begin SICMACT.TxtBuscarGeneral TxtBuscarGeneral 
         Height          =   375
         Left            =   1200
         TabIndex        =   14
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   0
         sPersOcupa      =   0
         sTitulo         =   ""
         lbUltimaInstancia=   0   'False
         PersPersoneria  =   0
         ColDesc         =   0
      End
      Begin VB.Label Label6 
         Caption         =   "Nombre"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblNombreSAF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Top             =   600
         Width           =   5895
      End
      Begin VB.Label Label4 
         Caption         =   "Código"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SICMACM"
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin SICMACT.TxtBuscar TxtBCodPers 
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   120
         Width           =   1860
         _ExtentX        =   3069
         _ExtentY        =   503
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.Label lblDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   960
         Width           =   5895
      End
      Begin VB.Label Label3 
         Caption         =   "Dirección"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   600
         Width           =   5895
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmUsuarioLeasing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsSql As String
Dim lsOption1() As String
Dim lsOption2() As String
Dim lsOption3() As String
Dim lsCodigo As String
Dim lsDescripcion As String

Private Sub CmdCancelar_Click()
    TxtBCodPers.Text = ""
    lblNombre.Caption = ""
    lblDireccion.Caption = ""
    TxtBuscarGeneral.Text = ""
    lblNombreSAF.Caption = ""
End Sub

Private Sub CmdGuardar_Click()
Dim obj As COMDCredito.DCOMleasing
Set obj = New COMDCredito.DCOMleasing
Call obj.ActualizarUsuarioSAFSICMACM(TxtBCodPers.Text, TxtBuscarGeneral.Text)
MsgBox "Datos se guardaron correctamente", vbApplicationModal, "Usuario SAF-SICMACM"
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
lsSql = "SELECT USU_CODIGO as cCodigo,USU_NOMBRE as cDescripcion "
lsSql = lsSql & " FROM SAF..SEG_USUARIO "

ReDim Preserve lsOption1(1 To 3)
lsOption1(1) = True
lsOption1(2) = "USU_CODIGO"
lsOption1(3) = "Usuario"

ReDim Preserve lsOption2(1 To 3)
lsOption2(1) = True
lsOption2(2) = "USU_NOMBRE"
lsOption2(3) = "Descripción"

ReDim Preserve lsOption3(1 To 3)
lsOption3(1) = False
lsOption3(2) = ""
lsOption3(3) = ""

Call TxtBuscarGeneral.inicio(lsSql, lsOption1, lsOption2, lsOption3, lsCodigo, lsDescripcion, "Busqueda de Usuarios Leasing")
CentraForm Me
End Sub

Private Sub TxtBCodPers_EmiteDatos()
lblNombre.Caption = TxtBCodPers.psDescripcion
lblDireccion.Caption = TxtBCodPers.sPersDireccion
End Sub

Private Sub TxtBuscarGeneral_EmiteDatos()
lblNombreSAF.Caption = TxtBuscarGeneral.psDescripcion
End Sub


