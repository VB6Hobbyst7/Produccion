VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCuentas 
   Caption         =   "Detalles de Operación"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9165
   Icon            =   "FrmCuentas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmCuentas.frx":030A
   ScaleHeight     =   2925
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView LstPlantilla 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Operacion"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Cod Concepto"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Desc Concepto"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cuenta Contable"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "OpeCtaDH"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Plantilla Contable"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "FrmCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oCon As COMConecta.DCOMConecta

Public Sub Inicio(ByVal psOpeCod As String)
    Set oCon = New COMConecta.DCOMConecta
    oCon.AbreConexion

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim sAge As String
    Dim lst As ListItem
    LstPlantilla.ListItems.Clear
    sql = " select cOpeCod,nConcepto,isnull(cDescripcion,''),cCtaContCod,cOpeCtaDH " & _
          " from dbo.OpeCtaNeg O " & _
          " left Join ProductoConcepto P on O.nConcepto=P.nPrdConceptoCod " & _
          " Where cOpeCod = '" & psOpeCod & "'"
    Set rs = oCon.CargaRecordSet(sql)
    If Not rs.EOF And Not rs.BOF Then
        Do Until rs.EOF
            Set lst = LstPlantilla.ListItems.Add(, , rs(0))
            lst.SubItems(1) = rs(1)
            lst.SubItems(2) = rs(2)
            lst.SubItems(3) = rs(3)
            lst.SubItems(4) = rs(4)
            rs.MoveNext
        Loop
    End If
    Me.Show 1
End Sub


