VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmVerCreditoSolWeb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Solictudes de Facilidades de Reprogramación"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Solicitudes de Reprogramación"
      Height          =   5205
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   10185
      Begin VB.CommandButton btnSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   8640
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtBuscar 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   5655
      End
      Begin MSComctlLib.ListView Lst 
         Height          =   3900
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   9945
         _ExtentX        =   17542
         _ExtentY        =   6879
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "N° Cuenta"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cliente"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Fecha Sol."
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label lblCliente 
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmVerCreditoSolWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lbLeasing As Boolean
Dim fbHistorial As Boolean
Dim fbOtros As Boolean
Dim CtlCtaCod As ActXCodCta
'JOEP20210206 Garantia covid
Dim gnTpAcceso As Integer
'JOEP20210206 Garantia covid

'Public Sub Inicio(Optional ByVal pCtlCtaCod As ActXCodCta, Optional ByVal psNombre = "")  'add jhcu 16-09-2020 add pnAdicional
Public Sub Inicio(Optional ByVal pCtlCtaCod As ActXCodCta, Optional ByVal psNombre As String = "", Optional ByVal pnOpcion As Integer = -1) 'Add JOEP20210206 garantia covid
    Dim rs As ADODB.Recordset
    Dim oVisualizacion As COMNCredito.NCOMVisualizacion
    Dim Item As ListItem
    gnTpAcceso = pnOpcion 'Add JOEP20210206 garantia covid
    
    Set oVisualizacion = New COMNCredito.NCOMVisualizacion
    'Set rs = oVisualizacion.VerSolWeb(gsCodAge, gsCodUser, psNombre)
    Set rs = oVisualizacion.VerSolWeb(gsCodAge, gsCodUser, psNombre, pnOpcion) 'Add JOEP20210206 garantia covid
    Set oVisualizacion = Nothing
    Lst.ListItems.Clear
    Do Until rs.EOF
        Set Item = Lst.ListItems.Add(, , rs!cCtaCod)
        Item.SubItems(1) = rs!Cliente
        Item.SubItems(2) = CStr(rs!dRegistro)
        rs.MoveNext
    Loop
    Set CtlCtaCod = pCtlCtaCod
    Set rs = Nothing
    Me.Show vbModal
'    Lst.SetFocus
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub btnSalir_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
    Lst.SetFocus
End Sub

Private Sub Lst_DblClick()

 If Not Lst.SelectedItem Is Nothing Then
   
        CtlCtaCod.NroCuenta = Lst.ListItems(Lst.SelectedItem.Index).Text
        CtlCtaCod.SetFocusCuenta
                Unload Me
  End If
   
End Sub

Private Sub Lst_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Lst_DblClick
    End If
End Sub

Private Sub txtBuscar_Change()
    Dim rs As ADODB.Recordset
    Dim oVisualizacion As COMNCredito.NCOMVisualizacion
    Dim Item As ListItem
    
    Set oVisualizacion = New COMNCredito.NCOMVisualizacion
    'Set rs = oVisualizacion.VerSolWeb(gsCodAge, gsCodUser, txtBuscar.Text)
    Set rs = oVisualizacion.VerSolWeb(gsCodAge, gsCodUser, txtBuscar.Text, gnTpAcceso) 'Add JOEP20210206 garantia covid
    Set oVisualizacion = Nothing
    Lst.ListItems.Clear
    Do Until rs.EOF
        Set Item = Lst.ListItems.Add(, , rs!cCtaCod)
        Item.SubItems(1) = rs!Cliente
        Item.SubItems(2) = CStr(rs!dRegistro)
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub
