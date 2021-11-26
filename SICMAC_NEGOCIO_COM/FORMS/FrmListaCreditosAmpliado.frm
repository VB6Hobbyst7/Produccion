VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmListaCreditosAmpliado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Creditos Ampliados"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   3705
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   7305
      Begin VB.Frame Frame2 
         Height          =   705
         Left            =   60
         TabIndex        =   6
         Top             =   2880
         Width           =   7005
         Begin VB.CommandButton CmdSalir 
            Caption         =   "&Salir"
            Height          =   375
            Left            =   5700
            TabIndex        =   9
            Top             =   210
            Width           =   1155
         End
         Begin VB.CommandButton CmdCancelar 
            Caption         =   "&Cancelar"
            Height          =   375
            Left            =   1350
            TabIndex        =   8
            Top             =   240
            Width           =   1155
         End
         Begin VB.CommandButton CmdAceptar 
            Caption         =   "&Aceptar"
            Height          =   375
            Left            =   90
            TabIndex        =   7
            Top             =   240
            Width           =   1155
         End
      End
      Begin MSDataGridLib.DataGrid DG 
         Height          =   2025
         Left            =   60
         TabIndex        =   5
         Top             =   750
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   3572
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   4
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1140
         TabIndex        =   3
         Top             =   240
         Width           =   4275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   2
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   30
         Left            =   600
         TabIndex        =   1
         Top             =   360
         Width           =   30
      End
   End
End
Attribute VB_Name = "FrmListaCreditosAmpliado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLista As ADODB.Recordset
Dim sCtaCod As String

Private Sub CmdAceptar_Click()
    If sCtaCod <> "" Then
        FrmGraAmpliado.CargarCredito (sCtaCod)
        Unload Me
    Else
        MsgBox "Debe seleccionar un credito", vbInformation, "AVISO"
    End If
End Sub

Private Sub CmdBuscar_Click()
    Dim rs As ADODB.Recordset
    Dim oAmpliacion As COMDCredito.DCOMAmpliacion
    Dim Item As ListItem
    
    ConfigurarGrid
    
    Set oAmpliacion = New COMDCredito.DCOMAmpliacion
    Set rs = oAmpliacion.ListaCreditosAmpliadosByNombre(txtNombre.Text)
    Set oAmpliacion = Nothing
    
    Do Until rs.EOF
        rsLista.AddNew
        rsLista(0) = rs!cCtaCod
        rsLista(1) = rs!cPersNombre
        rsLista.Update
        rs.MoveNext
    Loop
    DG.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Sub ConfigurarGrid()
    Set rsLista = New ADODB.Recordset
    
    With rsLista.Fields
        .Append "Cuenta", adVarChar, 18
        .Append "Nombre", adChar, 80
    End With
    
    rsLista.Open
    
    Set DG.DataSource = rsLista
    DG.AllowAddNew = False
    DG.AllowDelete = False
    DG.AllowUpdate = False
    
    DG.Columns(1).Width = 3000
End Sub

Private Sub DG_Click()
    Dim nCol As Integer
    
    nCol = DG.Col
    sCtaCod = rsLista(0)
End Sub

Private Sub Form_Load()
    ConfigurarGrid
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    Dim c  As String
     
     c = Chr(KeyAscii)
     c = UCase(c)
     KeyAscii = Asc(c)
     
     If KeyAscii = 13 Then
        CmdBuscar_Click
     End If
End Sub
