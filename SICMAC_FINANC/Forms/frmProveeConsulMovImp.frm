VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProveeConsulMovImp 
   Caption         =   "Selección de Impuestos"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6825
   Icon            =   "frmProveeConsulMovImp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "I m p u e s t o s"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   120
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   180
      Width           =   6615
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   3330
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   3330
      Width           =   1275
   End
   Begin MSComctlLib.ListView lvImp 
      Height          =   2355
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4154
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   7761
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Abreviatura"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmProveeConsulMovImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sSql As String
Dim rs   As ADODB.Recordset
Dim pbOk As Boolean
Dim oCon As DConecta

Private Sub cmdAceptar_Click()
pbOk = True
Me.Hide
End Sub

Private Sub cmdCancelar_Click()
pbOk = False
Me.Hide
End Sub

Private Sub Form_Load()
Dim lvItem As ListItem
CentraForm Me
Set oCon = New DConecta
oCon.AbreConexion
sSql = "SELECT i.cCtaContCod, c.cCtaContDesc, cImpAbrev " _
     & "FROM Impuesto i JOIN CtaCont c ON c.cCtaContCod = i.cCtaContCod "
Set rs = oCon.CargaRecordSet(sSql)
Do While Not rs.EOF
   Set lvItem = lvImp.ListItems.Add(, , rs!cCtaContCod)
   lvItem.SubItems(1) = rs!cCtaContDesc
   lvItem.SubItems(2) = rs!cImpAbrev
   rs.MoveNext
Loop
RSClose rs
oCon.CierraConexion
Set oCon = Nothing
End Sub

Public Property Get lOk() As Boolean
lOk = pbOk
End Property

Public Property Let lOk(ByVal vNewValue As Boolean)
pbOk = lOk
End Property
