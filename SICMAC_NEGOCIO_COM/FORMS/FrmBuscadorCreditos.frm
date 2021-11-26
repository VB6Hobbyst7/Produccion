VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBuscadorCreditos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscador Creditos"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   30
      TabIndex        =   1
      Top             =   2700
      Width           =   5805
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   255
         Left            =   90
         TabIndex        =   3
         Top             =   240
         Width           =   1005
      End
      Begin VB.CommandButton CmdSeleccionar 
         Caption         =   "Seleccionar"
         Height          =   255
         Left            =   1140
         TabIndex        =   2
         Top             =   240
         Width           =   1005
      End
   End
   Begin MSComctlLib.ListView LstwLista 
      Height          =   2640
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   4657
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Crédito"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Persona"
         Object.Width           =   7832
      EndProperty
   End
End
Attribute VB_Name = "FrmBuscadorCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cPersCod As String

Private Sub CmdBuscar_Click()
    Dim oAmpliado  As COMDCredito.DCOMAmpliacion
    Dim rs As ADODB.Recordset
    Dim L As ListItem
    
    
    LstwLista.ListItems.Clear
    Set oAmpliado = New COMDCredito.DCOMAmpliacion
    Set rs = oAmpliado.ListaCreditoPorAmpliar(cPersCod)
    Set oAmpliado = Nothing
    
    
    Do While Not rs.EOF
        
        Set L = LstwLista.ListItems.Add(, rs!cCtaCod, rs!cCtaCod)
        L.SubItems(1) = rs!cPersNombre
        rs.MoveNext
    Loop
    
    Set rs = Nothing
End Sub

Private Sub CmdSeleccionar_Click()
    Dim cCtaCod As String
    
    If Not LstwLista.SelectedItem Is Nothing Then
        cCtaCod = Mid(LstwLista.SelectedItem.Key, 1, Len(LstwLista.SelectedItem.Key))
        FrmCredAmpliado.cCtaCod = cCtaCod
        Unload Me
    Else
       MsgBox "Debe seleccionar un credito", vbInformation, "AVISO"
    End If
    
End Sub

Private Sub Form_Load()
    CmdBuscar_Click
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    Dim c As String
       c = Chr(KeyAscii)
       c = UCase(c)
       KeyAscii = Asc(c)
    If KeyAscii = 13 Then
       CmdBuscar_Click
    End If
End Sub
