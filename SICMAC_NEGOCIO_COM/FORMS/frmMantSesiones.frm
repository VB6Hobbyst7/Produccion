VERSION 5.00
Begin VB.Form frmMantSesiones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Sesiones de Usuarios"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   Icon            =   "frmMantSesiones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6000
      TabIndex        =   1
      Top             =   3720
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin SICMACT.FlexEdit feSesiones 
         Height          =   3255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   5741
         Cols0           =   8
         HighLight       =   1
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Usuario-Máquina-Fec. Ingreso-Aplicación-Activo-cPersCod-nTipoApp"
         EncabezadosAnchos=   "400-1000-1300-1200-1900-650-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-5-X-X"
         ListaControles  =   "0-0-0-0-0-4-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "R-C-C-C-L-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmMantSesiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************
'** Nombre : frmMantSesiones
'** Descripción : Formulario para administrar las sesiones de los usuarios en el Sicmac Negocio
'** Creación : JUEZ, 20160125 09:00:00 AM
'************************************************************************************************

Option Explicit

Dim oPista As COMManejador.Pista

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub feSesiones_Click()
    If feSesiones.Col = 5 Then
        If feSesiones.TextMatrix(feSesiones.row, 0) <> "" Then
            If feSesiones.TextMatrix(feSesiones.row, feSesiones.Col) = "." Then
                Set oPista = New COMManejador.Pista
                    oPista.ActualizarPistaSesion feSesiones.TextMatrix(feSesiones.row, 6), feSesiones.TextMatrix(feSesiones.row, 2), feSesiones.TextMatrix(feSesiones.row, 7)
                Set oPista = Nothing
                MsgBox "Sesión del usuario " & feSesiones.TextMatrix(feSesiones.row, 1) & " desactivada", vbInformation, "Aviso"
                CargarSesionesActivas
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    CargarSesionesActivas
End Sub

Private Sub CargarSesionesActivas()
Dim rs As ADODB.Recordset
    Set oPista = New COMManejador.Pista
        Set rs = oPista.ListarPistaSesiones(True)
    Set oPista = Nothing
        
    LimpiaFlex feSesiones
    If Not rs.EOF And Not rs.BOF Then
        Set feSesiones.Recordset = rs
    End If
    Set rs = Nothing
End Sub
