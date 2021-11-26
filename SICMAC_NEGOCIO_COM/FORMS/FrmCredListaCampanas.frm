VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmCredListaCampanas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Campañas"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6315
   Icon            =   "FrmCredListaCampanas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4155
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   6255
      Begin VB.Frame Frame2 
         Height          =   675
         Left            =   150
         TabIndex        =   2
         Top             =   3390
         Width           =   6045
         Begin VB.CommandButton CmdSeleccion 
            Caption         =   "S&eleccion"
            Height          =   345
            Left            =   150
            TabIndex        =   4
            Top             =   240
            Width           =   915
         End
         Begin VB.CommandButton CmdSalir 
            Caption         =   "&Salir"
            Height          =   345
            Left            =   4950
            TabIndex        =   3
            Top             =   210
            Width           =   915
         End
      End
      Begin MSComctlLib.ListView Lst 
         Height          =   3240
         Left            =   60
         TabIndex        =   1
         Top             =   150
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   5715
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "IdCampana"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Campaña"
            Object.Width           =   7937
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmCredListaCampanas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub Inicio(ByVal psAgenciaCod As String, Optional ByVal psCtaCod As String)
    Dim odCamp As COMDCredito.DCOMCampanas
    Dim rs As ADODB.Recordset
    Dim Item As ListItem
    Dim nIdCampana As Integer
    Dim i As Integer
    
    Set odCamp = New COMDCredito.DCOMCampanas
    
    Call odCamp.Cargar_Objetos_Controles(nIdCampana, psCtaCod, rs, psAgenciaCod)
    
    'If psCtaCod <> "" Then
    '    nIdCampana = odCamp.ObtenerIdCampanaxCuenta(psCtaCod)
    'End If
    'Set rs = odCamp.ListaCampanasXAgencia(psAgenciaCod)
    Set odCamp = Nothing
    
    Do Until rs.EOF
        Set Item = Lst.ListItems.Add(, , rs!IdCampana)
        Item.SubItems(1) = rs!cDescripcion
        rs.MoveNext
    Loop
    
    If nIdCampana <> -1 And nIdCampana > 0 Then
        For i = 1 To Lst.ListItems.Count
            If Lst.ListItems(i) = nIdCampana Then
                Lst.ListItems(i).Selected = True
                Exit For
            End If
        Next i
    End If
    
    Me.Show vbModal
End Sub


Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdSeleccion_Click()
    If Not Lst.SelectedItem Then
        frmCredSolicitud.nCampanaCod = Lst.ListItems(Lst.SelectedItem.Index)
        frmCredSolicitud.cCampanaDesc = Lst.ListItems(Lst.SelectedItem.Index).SubItems(1)
        Unload Me
    Else
        MsgBox "Debe seleccionar una campaña", vbInformation, "Aviso"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Lst.ListItems.Count <> 0 Then
        If Not Lst.SelectedItem Then
            frmCredSolicitud.nCampanaCod = Lst.ListItems(Lst.SelectedItem.Index)
            frmCredSolicitud.cCampanaDesc = Lst.ListItems(Lst.SelectedItem.Index).SubItems(1)
        End If
    End If
End Sub
