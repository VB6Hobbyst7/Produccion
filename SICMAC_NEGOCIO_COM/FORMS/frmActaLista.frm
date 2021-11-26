VERSION 5.00
Begin VB.Form frmActaLista 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usuarios de Operaciones"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   Icon            =   "frmActaLista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3780
      TabIndex        =   1
      Top             =   3180
      Width           =   1110
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1965
      TabIndex        =   0
      Top             =   3180
      Width           =   1110
   End
   Begin VB.Frame FraUsuario 
      Height          =   3015
      Left            =   105
      TabIndex        =   2
      Top             =   75
      Width           =   7095
      Begin SICMACT.FlexEdit grdLista 
         Height          =   2625
         Left            =   60
         TabIndex        =   3
         Top             =   210
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   4630
         Cols0           =   4
         HighLight       =   1
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Usuario-Nombre-cPersCod"
         EncabezadosAnchos=   "350-800-5500-0"
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-C"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "#"
         SelectionMode   =   1
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame fraActa 
      Height          =   3015
      Left            =   105
      TabIndex        =   4
      Top             =   60
      Width           =   7095
      Begin SICMACT.FlexEdit grdActa 
         Height          =   2625
         Left            =   60
         TabIndex        =   5
         Top             =   270
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   4630
         Cols0           =   6
         HighLight       =   1
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Fecha-Portador-NroActa-FechaRegistro-CodPortador"
         EncabezadosAnchos=   "350-1200-3600-1600-0-0"
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-C"
         FormatosEdit    =   "0-0-0-0-0-0"
         TextArray0      =   "#"
         SelectionMode   =   1
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmActaLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TipoRep As Integer

Private Sub CmdAceptar_Click()
  If TipoRep = 1 Then
     frmActaBilletesFalsos.TxtNroActa.Text = Trim(grdActa.TextMatrix(grdActa.Row, 3))
     frmActaBilletesFalsos.lblFechaReg.Caption = Trim(grdActa.TextMatrix(grdActa.Row, 4))
     
    
  Else
     frmActaBilletesFalsos.LblUsuVerificador.Caption = Trim(grdLista.TextMatrix(grdLista.Row, 1)) & ": " & Trim(grdLista.TextMatrix(grdLista.Row, 2))
     frmActaBilletesFalsos.LblUsuVerificador.Tag = Trim(grdLista.TextMatrix(grdLista.Row, 3))
  End If
  Unload Me
End Sub

Private Sub cmdCancelar_Click()
 Unload Me
End Sub
Public Sub Inicia(ByVal ntipo As Integer)
  TipoRep = ntipo
  If TipoRep = 1 Then
    fraActa.Visible = True
    FraUsuario.Visible = False
  Else
    fraActa.Visible = False
    FraUsuario.Visible = True
  End If
  CargaData (TipoRep)
  Me.Show 1
End Sub


Private Sub CargaData(ByVal ntipo As Integer)
Dim rs As Recordset
Dim CLSSERV As NCapServicios
Set CLSSERV = New NCapServicios

  If ntipo = 2 Then
           Set rs = CLSSERV.GetInfoRP(gsCodAge)
           Set grdLista.Recordset = rs
           
  Else
           Set rs = CLSSERV.GetInfoacta(gsCodUser)
           Set grdActa.Recordset = rs
           
  End If
  
End Sub

