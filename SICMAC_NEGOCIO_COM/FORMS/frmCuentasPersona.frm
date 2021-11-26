VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCuentasPersona 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuentas de la Persona"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   Icon            =   "frmCuentasPersona.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4935
      TabIndex        =   4
      Top             =   3465
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6045
      TabIndex        =   3
      Top             =   3465
      Width           =   1000
   End
   Begin MSComctlLib.ListView Lst 
      Height          =   2745
      Left            =   105
      TabIndex        =   2
      Top             =   630
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   4842
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cuenta"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Relacion"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Estado"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Monto"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Dias Vencidos"
         Object.Width           =   2469
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Persona"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   105
      TabIndex        =   1
      Top             =   240
      Width           =   765
   End
   Begin VB.Label lblNombrePersona 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   885
      TabIndex        =   0
      Top             =   210
      Width           =   6150
   End
End
Attribute VB_Name = "frmCuentasPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsCuenta As New COMDPersona.UCOMProdPersona

Public Function Inicio(ByVal psPersNombre As String, ByVal prCtas As ADODB.Recordset) As COMDPersona.UCOMProdPersona
    
    Dim iTem As ListItem
    Me.lblNombrePersona.Caption = psPersNombre
    
    If Not (prCtas.EOF And prCtas.EOF) Then
        Lst.ListItems.Clear
        Do While Not prCtas.EOF
                        
            Set iTem = Lst.ListItems.Add(, , prCtas("cCtaCod"))
            iTem.SubItems(1) = prCtas("cRelacion")
            iTem.SubItems(2) = prCtas("cEstado")
            iTem.SubItems(3) = Format(prCtas("nMonto"), "#,##0.00")
            iTem.SubItems(4) = IIf(DateDiff("d", prCtas("dVenc"), gdFecSis) < 0, "0", DateDiff("d", prCtas("dVenc"), gdFecSis))
            
            prCtas.MoveNext
        Loop
    Else
        MsgBox "Persona no posee creditos con estas condiciones", vbInformation, "Aviso"
    End If

    Me.Show 1
    Set Inicio = clsCuenta
    Set clsCuenta = Nothing
    
End Function

Private Sub CmdAceptar_Click()

    Dim sCta, sProd As String
    
    sCta = ""
    sProd = ""
    
    If Lst.ListItems.Count > 0 Then
        sCta = Lst.ListItems(Lst.SelectedItem.Index).Text
        sProd = Lst.ListItems.iTem(Lst.SelectedItem.Index).SubItems(1)
    End If
    
    clsCuenta.CargaDatos sCta, sProd, ""
    Unload Me
    
End Sub

Private Sub cmdCancelar_Click()

    clsCuenta.CargaDatos "", "", ""
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    Me.Caption = "Cuentas de la Persona"
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    Set clsCuenta = New COMDPersona.UCOMProdPersona
    
End Sub

