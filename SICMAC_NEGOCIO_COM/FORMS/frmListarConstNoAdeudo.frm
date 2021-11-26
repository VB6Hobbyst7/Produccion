VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form frmListarConstNoAdeudo 
   Caption         =   "Listar Constancia de No Adeudo"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9810
   Icon            =   "frmListarConstNoAdeudo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   9810
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   3855
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   9400
      Begin VB.CommandButton cmdEntregar 
         Caption         =   "&Entregar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         TabIndex        =   15
         Top             =   3360
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid dgBuscar 
         Height          =   3015
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   5318
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   1
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "iConstNoAdeudoId"
            Caption         =   "iConstNoAdeudoId"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "cPersCod"
            Caption         =   "cPersCod"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "cPersNombre"
            Caption         =   "Nombre"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "cFSolicitud"
            Caption         =   "F. Solicitud"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "cFEntrega"
            Caption         =   "F. Entrega"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "tMotivo"
            Caption         =   "Motivo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Estado"
            Caption         =   "Estado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            Size            =   182
            BeginProperty Column00 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1200.189
            EndProperty
         EndProperty
      End
      Begin VB.Label lblMensaje 
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   14
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros de Búsqueda:"
      Height          =   2415
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   8175
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   13
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   7455
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   1080
            TabIndex        =   11
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Format          =   70713345
            CurrentDate     =   39681
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   4200
            TabIndex        =   12
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Format          =   70713345
            CurrentDate     =   39681
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta:"
            Height          =   255
            Left            =   3360
            TabIndex        =   10
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Desde:"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.CheckBox chkFSolicitud 
         Alignment       =   1  'Right Justify
         Caption         =   "F. Solicitud"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1275
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   4320
         TabIndex        =   5
         Top             =   360
         Width           =   3375
      End
      Begin SICMACT.TxtBuscar txtCodigo 
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Top             =   360
         Width           =   1815
         _ExtentX        =   2990
         _ExtentY        =   556
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   3480
         TabIndex        =   3
         Top             =   360
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmListarConstNoAdeudo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsmensaje As String

Private Sub chkFSolicitud_Click()
    If chkFSolicitud.value = 1 Then
        Frame3.Visible = True
        Label1.Visible = True
        Label2.Visible = True
        DTPicker1.Visible = True
        DTPicker2.Visible = True
        DTPicker1.value = Date
        DTPicker2.value = Date
    Else
        Frame3.Visible = False
        Label1.Visible = False
        Label2.Visible = False
        DTPicker1.Visible = False
        DTPicker2.Visible = False
    End If
End Sub

Private Sub cmdBuscar_Click()
    BuscarDatos
End Sub

Private Sub cmdEntregar_Click()
    If lsmensaje = "" Then
        If ValidarEntrega = False Then
            frmEntregaConstNoAdeudo.Inicio dgBuscar.Columns(0).Text, dgBuscar.Columns(1).Text
            frmEntregaConstNoAdeudo.Show 1
        Else
            MsgBox "La Constancia ya fue entregada", vbInformation, "Aviso"
        End If
    End If
End Sub

Private Function ValidarEntrega() As Boolean

    Dim objCOMNCredito As COMNCredito.NCOMCredito
    Set objCOMNCredito = New COMNCredito.NCOMCredito
    Dim rs As ADODB.Recordset
    
    Set rs = objCOMNCredito.ValidarEntrega(dgBuscar.Columns(0).Text)
    
    If rs.RecordCount <> 0 Then
        ValidarEntrega = True
    Else
        ValidarEntrega = False
    End If

End Function

Public Sub BuscarDatos()
    Dim objCOMNCredito As COMNCredito.NCOMCredito
    Set objCOMNCredito = New COMNCredito.NCOMCredito
    Dim rs As Recordset
    lsmensaje = ""
    
    Set rs = objCOMNCredito.ObtenerListaConstNoAdeudo(txtCodigo.Text, IIf(chkFSolicitud.value = 0, "", DTPicker1.value), IIf(chkFSolicitud.value = 0, "", DTPicker2.value), lsmensaje)
    
    If lsmensaje = "" Then
        lblMensaje.Visible = False
        dgBuscar.Visible = True
        Set dgBuscar.DataSource = rs
        dgBuscar.Refresh
        Screen.MousePointer = 0
        dgBuscar.SetFocus
    Else
        Set dgBuscar.DataSource = Nothing
        dgBuscar.Refresh
        lblMensaje.Visible = True
        lblMensaje.Caption = "No Existen Datos"
        dgBuscar.Visible = False
    End If
    
    Set rs = Nothing
    Set objCOMNCredito = Nothing

End Sub

Private Sub Form_Load()
    Frame3.Visible = False
    Label1.Visible = False
    Label2.Visible = False
    DTPicker1.Visible = False
    DTPicker2.Visible = False
End Sub

Private Sub txtCodigo_EmiteDatos()
    Dim sCodigo As String
    Dim odRFa As COMDCredito.DCOMRFA
    Dim rs As ADODB.Recordset
    
    If txtCodigo.Text <> "" Then
        sCodigo = txtCodigo.Text
        Set odRFa = New COMDCredito.DCOMRFA
        Set rs = odRFa.BuscarPersona(sCodigo)
        Set odRFa = Nothing
    
        If Not rs.EOF And Not rs.BOF Then
            txtNombre.Text = rs!cPersNombre
        End If
        Set rs = Nothing
    End If
    
End Sub
